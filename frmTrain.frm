VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTrain 
   BackColor       =   &H8000000E&
   Caption         =   "Train"
   ClientHeight    =   8145
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   17775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   17775
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   0
      Top             =   2160
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   8400
      Picture         =   "frmTrain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   6600
      Picture         =   "frmTrain.frx":28FE
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   10200
      Picture         =   "frmTrain.frx":5178
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   4800
      Picture         =   "frmTrain.frx":7A76
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "User Input Information"
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
      Height          =   6135
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   15975
      Begin VB.TextBox txtUTrainNo 
         BackColor       =   &H8000000F&
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
         Left            =   2160
         TabIndex        =   36
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chkSat1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5760
         MaskColor       =   &H000000C0&
         TabIndex        =   34
         Top             =   5040
         Width           =   615
      End
      Begin VB.CheckBox chkFri1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5160
         MaskColor       =   &H000000C0&
         TabIndex        =   33
         Top             =   5040
         Width           =   615
      End
      Begin VB.CheckBox chkThr1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Thr"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4440
         MaskColor       =   &H000000C0&
         TabIndex        =   32
         Top             =   5040
         Width           =   735
      End
      Begin VB.CheckBox chkWed1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3600
         MaskColor       =   &H000000C0&
         TabIndex        =   31
         Top             =   5040
         Width           =   855
      End
      Begin VB.CheckBox chkTue1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2880
         MaskColor       =   &H000000C0&
         TabIndex        =   30
         Top             =   5040
         Width           =   735
      End
      Begin VB.CheckBox chkMon1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2160
         MaskColor       =   &H000000C0&
         TabIndex        =   29
         Top             =   5040
         Width           =   735
      End
      Begin VB.CheckBox chkSun1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6360
         MaskColor       =   &H000000C0&
         TabIndex        =   28
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   5400
         Picture         =   "frmTrain.frx":A1B8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkSun 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6360
         MaskColor       =   &H000000C0&
         TabIndex        =   17
         Top             =   4200
         Width           =   615
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "Route Details"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   5655
         Left            =   7080
         TabIndex        =   14
         Top             =   240
         Width           =   8655
         Begin VB.ComboBox cmbRoute 
            BackColor       =   &H8000000F&
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
            Left            =   2880
            TabIndex        =   15
            Top             =   360
            Width           =   2895
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   4575
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   8070
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BackColor       =   -2147483634
            ForeColor       =   192
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "UP Route"
            TabPicture(0)   =   "frmTrain.frx":BABA
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Picture1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "flexGridRoute"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "MaskEdBox1"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Text1"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "DOWN Route"
            TabPicture(1)   =   "frmTrain.frx":BAD6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Picture2"
            Tab(1).Control(1)=   "flexGridRoute1"
            Tab(1).Control(2)=   "MaskEdBox2"
            Tab(1).Control(3)=   "Text2"
            Tab(1).ControlCount=   4
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "Bodoni MT"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   -73320
               TabIndex        =   38
               Text            =   "Text1"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Bodoni MT"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               TabIndex        =   25
               Text            =   "Text1"
               Top             =   480
               Width           =   975
            End
            Begin MSMask.MaskEdBox MaskEdBox1 
               Height          =   375
               Left            =   480
               TabIndex        =   26
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
               _Version        =   393216
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Bodoni MT"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSFlexGridLib.MSFlexGrid flexGridRoute 
               Height          =   3975
               Left            =   480
               TabIndex        =   27
               Top             =   480
               Width           =   7455
               _ExtentX        =   13150
               _ExtentY        =   7011
               _Version        =   393216
               BackColor       =   -2147483634
               ForeColor       =   192
               BackColorSel    =   192
               GridColor       =   255
               GridColorFixed  =   192
               GridLinesFixed  =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Bodoni MT"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSMask.MaskEdBox MaskEdBox2 
               Height          =   375
               Left            =   -74520
               TabIndex        =   39
               Top             =   480
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   661
               _Version        =   393216
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Bodoni MT"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "hh:mm"
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSFlexGridLib.MSFlexGrid flexGridRoute1 
               Height          =   3975
               Left            =   -74520
               TabIndex        =   40
               Top             =   480
               Width           =   7455
               _ExtentX        =   13150
               _ExtentY        =   7011
               _Version        =   393216
               BackColor       =   -2147483634
               ForeColor       =   192
               BackColorSel    =   192
               GridColor       =   255
               GridColorFixed  =   192
               GridLinesFixed  =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Bodoni MT"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.PictureBox Picture1 
               Height          =   4335
               Left            =   0
               Picture         =   "frmTrain.frx":BAF2
               ScaleHeight     =   4275
               ScaleWidth      =   8355
               TabIndex        =   41
               Top             =   360
               Width           =   8415
            End
            Begin VB.PictureBox Picture2 
               Height          =   4335
               Left            =   -75000
               Picture         =   "frmTrain.frx":34AE34
               ScaleHeight     =   4275
               ScaleWidth      =   8355
               TabIndex        =   42
               Top             =   360
               Width           =   8415
            End
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            Caption         =   "Route Name:"
            BeginProperty Font 
               Name            =   "Bodoni MT"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.CheckBox chkMon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2160
         MaskColor       =   &H000000C0&
         TabIndex        =   13
         Top             =   4200
         Width           =   735
      End
      Begin VB.CheckBox chkTue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2880
         MaskColor       =   &H000000C0&
         TabIndex        =   12
         Top             =   4200
         Width           =   735
      End
      Begin VB.CheckBox chkWed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3600
         MaskColor       =   &H000000C0&
         TabIndex        =   11
         Top             =   4200
         Width           =   855
      End
      Begin VB.CheckBox chkThr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Thr"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4440
         MaskColor       =   &H000000C0&
         TabIndex        =   10
         Top             =   4200
         Width           =   735
      End
      Begin VB.CheckBox chkFri 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5160
         MaskColor       =   &H000000C0&
         TabIndex        =   9
         Top             =   4200
         Width           =   615
      End
      Begin VB.CheckBox chkSat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5760
         MaskColor       =   &H000000C0&
         TabIndex        =   8
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtTrainName 
         BackColor       =   &H8000000F&
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
         Left            =   2160
         TabIndex        =   6
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtDTrainNo 
         BackColor       =   &H8000000F&
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
         Left            =   2160
         TabIndex        =   5
         Top             =   2280
         Width           =   2895
      End
      Begin VB.ComboBox cmbTrainType 
         BackColor       =   &H8000000F&
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
         ItemData        =   "frmTrain.frx":68A176
         Left            =   2160
         List            =   "frmTrain.frx":68A178
         TabIndex        =   1
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Up Train No.:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   37
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Down Train Days:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Up Train Days:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Train Type:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Train Name:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Down Train No.:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   2400
         Width           =   1935
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3600
      TabIndex        =   43
      Top             =   600
      Width           =   12615
   End
   Begin VB.Image Image7 
      Height          =   525
      Left            =   2040
      Picture         =   "frmTrain.frx":68A17A
      Top             =   480
      Width           =   14565
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   8280
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   16920
      Picture         =   "frmTrain.frx":6A3068
      Top             =   960
      Width           =   825
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   840
      Picture         =   "frmTrain.frx":6B3682
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   5985
      Left            =   120
      Picture         =   "frmTrain.frx":6B3FCD
      Top             =   960
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11280
      Picture         =   "frmTrain.frx":6C45E7
      Top             =   0
      Width           =   11535
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD Train INFORMATION"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   1560
      TabIndex        =   19
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   -120
      Picture         =   "frmTrain.frx":6C4AC1
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmTrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdTrain As ADODB.Command
Dim rstTrainType As ADODB.Recordset
Dim rstRoute As ADODB.Recordset
Dim rstTrain As ADODB.Recordset
Dim ID As Integer
Dim f As Integer
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
    If saveUpdate = 2 Then
        If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
            Set cmdTrain = New ADODB.Command
            cmdTrain.CommandType = adCmdText
            cmdTrain.ActiveConnection = railCn
            cmdTrain.CommandText = "delete from trainroute where trainno='" & frmTrainDialog.upTrainNo & "' or trainno='" & frmTrainDialog.downTrainNo & "'"
            cmdTrain.Execute
            cmdTrain.CommandText = "delete from days where trainno='" & frmTrainDialog.upTrainNo & "' or trainno='" & frmTrainDialog.downTrainNo & "'"
            cmdTrain.Execute
            cmdTrain.CommandText = "delete from train where uptrainno='" & frmTrainDialog.upTrainNo & "'"
            cmdTrain.Execute
            MsgBox "Record Successfully Deleted", vbInformation
            txtUTrainNo.Enabled = False
            txtDTrainNo.Enabled = False
            txtTrainName.Enabled = False
            cmbTrainType.Enabled = False
            cmbRoute.Enabled = False
            chkMon.Enabled = False
            chkTue.Enabled = False
            chkWed.Enabled = False
            chkThr.Enabled = False
            chkFri.Enabled = False
            chkSat.Enabled = False
            chkSun.Enabled = False
            chkMon.Value = 0
            chkTue.Value = 0
            chkWed.Value = 0
            chkThr.Value = 0
            chkFri.Value = 0
            chkSat.Value = 0
            chkSun.Value = 0
            chkMon1.Enabled = False
            chkTue1.Enabled = False
            chkWed1.Enabled = False
            chkThr1.Enabled = False
            chkFri1.Enabled = False
            chkSat1.Enabled = False
            chkSun1.Enabled = False
            chkMon1.Value = 0
            chkTue1.Value = 0
            chkWed1.Value = 0
            chkThr1.Value = 0
            chkFri1.Value = 0
            chkSat1.Value = 0
            chkSun1.Value = 0
            txtUTrainNo.BackColor = vbButtonFace
            txtDTrainNo.BackColor = vbButtonFace
            txtTrainName.BackColor = vbButtonFace
            cmbTrainType.BackColor = vbButtonFace
            cmbRoute.BackColor = vbButtonFace
            txtUTrainNo.Text = ""
            txtDTrainNo.Text = ""
            txtTrainName.Text = ""
            cmbTrainType.ListIndex = -1
            cmbRoute.ListIndex = -1
            flexGridRoute.Rows = 1
            flexGridRoute.Cols = 7
            MaskEdBox1.Visible = False
            Text1.Visible = False
            flexGridRoute1.Rows = 1
            flexGridRoute1.Cols = 7
            MaskEdBox2.Visible = False
            Text2.Visible = False
            saveUpdate = 0
        End If
    Else
        MsgBox "Please Search and Select the Record", vbCritical
    End If
label:
Select Case Err.Number
   Case -2147467259
    MsgBox Err.Description, vbCritical
End Select
End Sub

Private Sub cmdNew_Click()
    txtUTrainNo.Enabled = True
    txtDTrainNo.Enabled = True
    txtTrainName.Enabled = True
    cmbTrainType.Enabled = True
    cmbRoute.Enabled = True
    chkMon.Enabled = True
    chkTue.Enabled = True
    chkWed.Enabled = True
    chkThr.Enabled = True
    chkFri.Enabled = True
    chkSat.Enabled = True
    chkSun.Enabled = True
    chkMon.Value = 0
    chkTue.Value = 0
    chkWed.Value = 0
    chkThr.Value = 0
    chkFri.Value = 0
    chkSat.Value = 0
    chkSun.Value = 0
    chkMon1.Enabled = True
    chkTue1.Enabled = True
    chkWed1.Enabled = True
    chkThr1.Enabled = True
    chkFri1.Enabled = True
    chkSat1.Enabled = True
    chkSun1.Enabled = True
    chkMon1.Value = 0
    chkTue1.Value = 0
    chkWed1.Value = 0
    chkThr1.Value = 0
    chkFri1.Value = 0
    chkSat1.Value = 0
    chkSun1.Value = 0
    txtUTrainNo.BackColor = vbHighlightText
    txtDTrainNo.BackColor = vbHighlightText
    txtTrainName.BackColor = vbHighlightText
    cmbTrainType.BackColor = vbHighlightText
    cmbRoute.BackColor = vbHighlightText
    saveUpdate = 1
    txtUTrainNo.Text = ""
    txtDTrainNo.Text = ""
    txtTrainName.Text = ""
    cmbTrainType.ListIndex = -1
    cmbRoute.ListIndex = -1
    flexGridRoute.Rows = 1
    flexGridRoute.Cols = 7
    flexGridRoute1.Rows = 1
    flexGridRoute1.Cols = 7
End Sub

Private Function checkDays() As Integer
    If (chkMon1.Value = 1 Or chkTue1.Value = 1 Or chkWed1.Value = 1 Or chkThr1.Value = 1 Or chkFri1.Value = 1 Or chkSat1.Value = 1 Or chkSun1.Value = 1) And (chkMon.Value = 1 Or chkTue.Value = 1 Or chkWed.Value = 1 Or chkThr.Value = 1 Or chkFri.Value = 1 Or chkSat.Value = 1 Or chkSun.Value = 1) Then
        checkDays = 1
    Else
        checkDays = 0
    End If
End Function

Private Function checkRoute() As Integer
    Dim flag As Integer
    Dim flag1 As Integer
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
            flag = flag + 1
        End If
    Next
    For i = 1 To flexGridRoute1.Rows - 1
        If flexGridRoute1.TextMatrix(i, 1) = strChecked Then
            flag1 = flag1 + 1
        End If
    Next
    If flag >= 2 And flag1 >= 2 Then
        checkRoute = 1
    Else
        checkRoute = 0
    End If
End Function
Private Function checkDaysSequencing1() As Boolean
    For i = 1 To flexGridRoute1.Rows - 1
        If flexGridRoute1.TextMatrix(i, 1) = strChecked Then
            If flexGridRoute1.TextMatrix(i, 6) = 1 Then
                checkDaysSequencing1 = True
                Exit Function
            Else
                checkDaysSequencing1 = False
                Exit Function
            End If
        End If
    Next
End Function
Private Function checkDaysSequencing() As Boolean
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
            If flexGridRoute.TextMatrix(i, 6) = 1 Then
                checkDaysSequencing = True
                Exit Function
            Else
                checkDaysSequencing = False
                Exit Function
            End If
        End If
    Next
End Function
Private Function compareTime(time1 As String, time2 As String) As Boolean
    If Mid(time1, 1, 2) > Mid(time2, 1, 2) Then
        compareTime = True
        Exit Function
    ElseIf Mid(time1, 1, 2) < Mid(time2, 1, 2) Then
        compareTime = False
        Exit Function
    ElseIf Mid(time1, 1, 2) = Mid(time2, 1, 2) And Mid(time1, 4, 2) > Mid(time2, 4, 2) Then
        compareTime = True
        Exit Function
    ElseIf Mid(time1, 1, 2) = Mid(time2, 1, 2) And Mid(time1, 4, 2) < Mid(time2, 4, 2) Then
        compareTime = True
        Exit Function
    ElseIf Mid(time1, 1, 2) = Mid(time2, 1, 2) And Mid(time1, 4, 2) = Mid(time2, 4, 2) Then
        compareTime = False
        Exit Function

    End If
End Function

Private Function selectTotal() As Integer
    Dim flag As Integer
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
            flag = flag + 1
        End If
    Next
    selectTotal = flag
End Function
Private Function CheckTime() As Boolean
    Dim preTime As String * 5
    Dim flag As Integer
    Dim tim As String * 5
    tim = "00:00"
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
            tim = addTime(tim, diffTime(flexGridRoute.TextMatrix(i, 5), flexGridRoute.TextMatrix(i, 4)))
        End If
    Next
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
            If flag = 1 Then
                tim = addTime(tim, diffTime(flexGridRoute.TextMatrix(i, 4), preTime))
            End If
            preTime = flexGridRoute.TextMatrix(i, 5)
            flag = 1
        End If
    Next
    If Val(Mid(tim, 1, 2)) >= 24 Then
    CheckTime = False
    Else
    CheckTime = True
    End If
End Function
Private Function CheckTime1() As Boolean
    Dim preTime As String * 5
    Dim flag As Integer
    Dim tim As String * 5
    tim = "00:00"
    For i = 1 To flexGridRoute1.Rows - 1
        If flexGridRoute1.TextMatrix(i, 1) = strChecked Then
            tim = addTime(tim, diffTime(flexGridRoute1.TextMatrix(i, 5), flexGridRoute1.TextMatrix(i, 4)))
        End If
    Next
    For i = 1 To flexGridRoute1.Rows - 1
        If flexGridRoute1.TextMatrix(i, 1) = strChecked Then
            If flag = 1 Then
                tim = addTime(tim, diffTime(flexGridRoute1.TextMatrix(i, 4), preTime))
            End If
            preTime = flexGridRoute1.TextMatrix(i, 5)
            flag = 1
        End If
    Next
    If Val(Mid(tim, 1, 2)) >= 24 Then
    CheckTime1 = False
    Else
    CheckTime1 = True
    End If
End Function


Private Function CheckTimeSequence() As Boolean
    f = 0
    Dim flag  As Integer
    Dim tim As String * 5
    Dim day As Integer
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
        f = f + 1
            tim = flexGridRoute.TextMatrix(i, 5)
            If Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) = Mid(flexGridRoute.TextMatrix(i, 5), 1, 2) And Mid(flexGridRoute.TextMatrix(i, 4), 4, 2) = Mid(flexGridRoute.TextMatrix(i, 5), 4, 2) Then
                CheckTimeSequence = False
                Exit Function
            ElseIf Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) > Mid(flexGridRoute.TextMatrix(i, 5), 1, 2) Or (Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) = Mid(flexGridRoute.TextMatrix(i, 5), 1, 2) And Mid(flexGridRoute.TextMatrix(i, 4), 4, 2) > Mid(flexGridRoute.TextMatrix(i, 5), 4, 2)) Then
                If flag < selectTotal() - 1 Then
                    For j = i + 1 To flexGridRoute.Rows - 1
                        If flexGridRoute.TextMatrix(j, 1) = strChecked Then
                            If (Mid(flexGridRoute.TextMatrix(j, 4), 1, 2) < tim Or (Mid(flexGridRoute.TextMatrix(j, 4), 1, 2) = tim And Mid(flexGridRoute.TextMatrix(j, 4), 4, 2) < tim)) Then
                                If flexGridRoute.TextMatrix(j, 6) <= day + 2 Then
                                    CheckTimeSequence = False
                                    Exit Function
                                Else
                                    Exit For
                                End If
                            Else
                                If flexGridRoute.TextMatrix(j, 6) <= day + 1 Then
                                    CheckTimeSequence = False
                                    Exit Function
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            ElseIf Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) < Mid(flexGridRoute.TextMatrix(i, 5), 1, 2) Or (Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) = Mid(flexGridRoute.TextMatrix(i, 5), 1, 2) And Mid(flexGridRoute.TextMatrix(i, 4), 4, 2) < Mid(flexGridRoute.TextMatrix(i, 5), 4, 2)) Then
                If flag < selectTotal() - 1 And f > 1 Then
                    For j = i + 1 To flexGridRoute.Rows - 1
                        If flexGridRoute.TextMatrix(j, 1) = strChecked Then
                            If (Mid(flexGridRoute.TextMatrix(j, 4), 1, 2) > tim Or (Mid(flexGridRoute.TextMatrix(j, 4), 1, 2) = tim And Mid(flexGridRoute.TextMatrix(j, 4), 4, 2) > tim)) Then
                                If flexGridRoute.TextMatrix(j, 6) > day + 1 Then
                                    CheckTimeSequence = False
                                    Exit Function
                                Else
                                    Exit For
                                End If
                            Else
                                If flexGridRoute.TextMatrix(j, 6) <> day + 1 Then
                                    CheckTimeSequence = False
                                    Exit Function
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                If flag > 0 Then
                    If (Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) > Mid(flexGridRoute.TextMatrix(i, 5), 1, 2) Or (Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) = Mid(flexGridRoute.TextMatrix(i, 5), 1, 2) And Mid(flexGridRoute.TextMatrix(i, 4), 4, 2) > Mid(flexGridRoute.TextMatrix(i, 5), 4, 2))) And (Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) > tim Or (Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) = tim And Mid(flexGridRoute.TextMatrix(i, 4), 4, 2) > tim)) And day <> Val(flexGridRoute.TextMatrix(i, 6)) Then
                        CheckTimeSequence = False
                        Exit Function
                    ElseIf Not (Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) > tim Or (Mid(flexGridRoute.TextMatrix(i, 4), 1, 2) = tim And Mid(flexGridRoute.TextMatrix(i, 4), 4, 2) > tim)) And day = Val(flexGridRoute.TextMatrix(i, 4)) Then
                        CheckTimeSequence = False
                        Exit Function
                    End If
                End If
            End If
            flag = flag + 1
            day = flexGridRoute.TextMatrix(i, 6)

        End If
    Next
    CheckTimeSequence = True
End Function
Private Function CheckTimeSequence1() As Boolean
    
    f = 0
    Dim flag  As Integer
    Dim tim As String * 5
    Dim day As Integer
    For i = 1 To flexGridRoute1.Rows - 1
        If flexGridRoute1.TextMatrix(i, 1) = strChecked Then
        f = f + 1
            tim = flexGridRoute1.TextMatrix(i, 5)
            If Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) = Mid(flexGridRoute1.TextMatrix(i, 5), 1, 2) And Mid(flexGridRoute1.TextMatrix(i, 4), 4, 2) = Mid(flexGridRoute1.TextMatrix(i, 5), 4, 2) Then
                CheckTimeSequence1 = False
                Exit Function
            ElseIf Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) > Mid(flexGridRoute1.TextMatrix(i, 5), 1, 2) Or (Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) = Mid(flexGridRoute1.TextMatrix(i, 5), 1, 2) And Mid(flexGridRoute1.TextMatrix(i, 4), 4, 2) > Mid(flexGridRoute1.TextMatrix(i, 5), 4, 2)) Then
                If flag < selectTotal() - 1 Then
                    For j = i + 1 To flexGridRoute1.Rows - 1
                        If flexGridRoute1.TextMatrix(j, 1) = strChecked Then
                            If (Mid(flexGridRoute1.TextMatrix(j, 4), 1, 2) < tim Or (Mid(flexGridRoute1.TextMatrix(j, 4), 1, 2) = tim And Mid(flexGridRoute1.TextMatrix(j, 4), 4, 2) < tim)) Then
                                If flexGridRoute1.TextMatrix(j, 6) <= day + 2 Then
                                    CheckTimeSequence1 = False
                                    Exit Function
                                Else
                                    Exit For
                                End If
                            Else
                                If flexGridRoute1.TextMatrix(j, 6) <= day + 1 Then
                                    CheckTimeSequence1 = False
                                    Exit Function
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            ElseIf Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) < Mid(flexGridRoute1.TextMatrix(i, 5), 1, 2) Or (Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) = Mid(flexGridRoute1.TextMatrix(i, 5), 1, 2) And Mid(flexGridRoute1.TextMatrix(i, 4), 4, 2) < Mid(flexGridRoute1.TextMatrix(i, 5), 4, 2)) Then
                If flag < selectTotal() - 1 And f > 1 Then
                    For j = i + 1 To flexGridRoute1.Rows - 1
                        If flexGridRoute1.TextMatrix(j, 1) = strChecked Then
                            If (Mid(flexGridRoute1.TextMatrix(j, 4), 1, 2) > tim Or (Mid(flexGridRoute1.TextMatrix(j, 4), 1, 2) = tim And Mid(flexGridRoute1.TextMatrix(j, 4), 4, 2) > tim)) Then
                                If flexGridRoute1.TextMatrix(j, 6) > day + 1 Then
                                    CheckTimeSequence1 = False
                                    Exit Function
                                Else
                                    Exit For
                                End If
                            Else
                                If flexGridRoute1.TextMatrix(j, 6) <> day + 1 Then
                                    CheckTimeSequence1 = False
                                    Exit Function
                                Else
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                If flag > 0 Then
                    If (Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) > Mid(flexGridRoute1.TextMatrix(i, 5), 1, 2) Or (Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) = Mid(flexGridRoute1.TextMatrix(i, 5), 1, 2) And Mid(flexGridRoute1.TextMatrix(i, 4), 4, 2) > Mid(flexGridRoute1.TextMatrix(i, 5), 4, 2))) And (Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) > tim Or (Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) = tim And Mid(flexGridRoute1.TextMatrix(i, 4), 4, 2) > tim)) And day <> Val(flexGridRoute1.TextMatrix(i, 6)) Then
                        CheckTimeSequence1 = False
                        Exit Function
                    ElseIf Not (Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) > tim Or (Mid(flexGridRoute1.TextMatrix(i, 4), 1, 2) = tim And Mid(flexGridRoute1.TextMatrix(i, 4), 4, 2) > tim)) And day = Val(flexGridRoute1.TextMatrix(i, 4)) Then
                        CheckTimeSequence1 = False
                        Exit Function
                    End If
                End If
            End If
            flag = flag + 1
            day = flexGridRoute1.TextMatrix(i, 6)

        End If
    Next
    CheckTimeSequence1 = True
End Function
Private Function checkStn() As Boolean
    Dim stn() As Variant
    Dim j As Integer
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strUnChecked Then
             ReDim Preserve stn(j + 1) As Variant
            stn(j) = flexGridRoute.TextMatrix(i, 2)
            j = j + 1
        End If
    Next
    For i = 1 To j
        For k = 1 To flexGridRoute1.Rows - 1
            If flexGridRoute1.TextMatrix(k, 2) = stn(i - 1) Then
                If flexGridRoute1.TextMatrix(k, 1) <> strUnChecked Then
                    checkStn = False
                    Exit Function
                End If
            End If
        Next
    Next
    checkStn = True
End Function
Private Function checkStn1() As Boolean
    Dim stn() As Variant
    Dim j As Integer
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
             ReDim Preserve stn(j + 1) As Variant
            stn(j) = flexGridRoute.TextMatrix(i, 2)
            j = j + 1
        End If
    Next
    For i = 1 To j
        For k = 1 To flexGridRoute1.Rows - 1
            If flexGridRoute1.TextMatrix(k, 2) = stn(i - 1) Then
                If flexGridRoute1.TextMatrix(k, 1) <> strChecked Then
                    checkStn1 = False
                    Exit Function
                End If
            End If
        Next
    Next
    checkStn1 = True
End Function

Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
    If Len(txtDTrainNo.Text) = 5 And Len(txtUTrainNo.Text) = 5 Then
        If checkStn() And checkStn1() And txtDTrainNo.Text <> "" And txtUTrainNo.Text <> "" And txtTrainName.Text <> "" And cmbTrainType.ListIndex <> -1 And cmbRoute.ListIndex <> -1 And checkEmptyFields() Then
            If checkDays() Then
                If checkRoute() Then
                    If checkDaysSequencing() And checkDaysSequencing1() Then
                        If CheckTime() And CheckTime1 Then
                            If CheckTimeSequence() And CheckTimeSequence1() Then
                                Set cmdTrain = New ADODB.Command
                                Set rstTrain = New ADODB.Recordset
                                cmdTrain.CommandType = adCmdText
                                cmdTrain.ActiveConnection = railCn
                                If saveUpdate = 1 Then
                                    cmdTrain.CommandText = "insert into train values('" & txtUTrainNo & "','" & txtDTrainNo & "'," & cmbRoute.ItemData(cmbRoute.ListIndex) & ",'" & txtTrainName & "'," & cmbTrainType.ItemData(cmbTrainType.ListIndex) & ")"
                                    cmdTrain.Execute
                                    cmdTrain.CommandText = "insert into days values('" & txtUTrainNo.Text & "'," & chkSun.Value & "," & chkMon.Value & "," & chkTue.Value & "," & chkWed.Value & "," & chkThr.Value & "," & chkFri.Value & "," & chkSat.Value & ")"
                                    cmdTrain.Execute
                                    Dim flag As Integer
                                    For i = 1 To flexGridRoute.Rows - 1
                                        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
                                            flag = flag + 1
                                            cmdTrain.CommandText = "insert into trainroute values('" & txtUTrainNo & "'," & flexGridRoute.TextMatrix(i, 2) & ",'" & flexGridRoute.TextMatrix(i, 4) & "','" & flexGridRoute.TextMatrix(i, 5) & "'," & flag & "," & flexGridRoute.TextMatrix(i, 6) & ")"
                                            cmdTrain.Execute
                                        End If
                                    Next
                                    cmdTrain.CommandText = "insert into days values('" & txtDTrainNo.Text & "'," & chkSun1.Value & "," & chkMon1.Value & "," & chkTue1.Value & "," & chkWed1.Value & "," & chkThr1.Value & "," & chkFri1.Value & "," & chkSat1.Value & ")"
                                    cmdTrain.Execute
                                    Dim flag1 As Integer
                                    For i = 1 To flexGridRoute1.Rows - 1
                                        If flexGridRoute1.TextMatrix(i, 1) = strChecked Then
                                            flag1 = flag1 + 1
                                            cmdTrain.CommandText = "insert into trainroute values('" & txtDTrainNo & "'," & flexGridRoute1.TextMatrix(i, 2) & ",'" & flexGridRoute1.TextMatrix(i, 4) & "','" & flexGridRoute1.TextMatrix(i, 5) & "'," & flag1 & "," & flexGridRoute1.TextMatrix(i, 6) & ")"
                                            cmdTrain.Execute
                                        End If
                                    Next
                                    saveUpdate = 0
                                    MsgBox "Record Successfully Saved", vbInformation
                                ElseIf saveUpdate = 2 Then
'                                    If frmTrainDialog.upTrainNo = Val(txtUTrainNo.Text) And frmTrainDialog.downTrainNo = Val(txtDTrainNo.Text) Then
'                                        cmdTrain.CommandText = "update train set trainName='" & txtTrainName.Text & "',routeid=" & cmbRoute.ItemData(cmbRoute.ListIndex) & ",traintypeid=" & cmbTrainType.ItemData(cmbTrainType.ListIndex) & " where trainno='" & frmTrainDialog.upTrainNo & "'"
'                                        cmdTrain.Execute
'                                        cmdTrain.CommandText = "update days set sun=" & chkSun.Value & ",mon=" & chkMon.Value & ",tue=" & chkTue.Value & ",wed=" & chkWed.Value & ",thru=" & chkThr.Value & ",fri=" & chkFri.Value & ",sat=" & chkSat.Value & " where trainno='" & frmTrainDialog.upTrainNo & "'"
'                                        cmdTrain.Execute
'                                        cmdTrain.CommandText = "delete from trainroute where trainno='" & frmTrainDialog.upTrainNo & "'"
'                                        cmdTrain.Execute
'                                        For i = 1 To flexGridRoute.Rows - 1
'                                            If flexGridRoute.TextMatrix(i, 1) = strChecked Then
'                                                flag = flag + 1
'                                                cmdTrain.CommandText = "insert into trainroute values('" & frmTrainDialog.upTrainNo & "'," & flexGridRoute.TextMatrix(i, 2) & ",'" & flexGridRoute.TextMatrix(i, 4) & "','" & flexGridRoute.TextMatrix(i, 5) & "'," & flag & "," & flexGridRoute.TextMatrix(i, 6) & ")"
'                                                cmdTrain.Execute
'                                            End If
'                                        Next
'                                    Else
                                        cmdTrain.CommandText = "update train set uptrainNo='" & txtUTrainNo.Text & "',downtrainNo='" & txtDTrainNo.Text & "',trainName='" & txtTrainName.Text & "',routeid=" & cmbRoute.ItemData(cmbRoute.ListIndex) & ",traintypeid=" & cmbTrainType.ItemData(cmbTrainType.ListIndex) & " where uptrainno='" & frmTrainDialog.upTrainNo & "'"
                                        cmdTrain.Execute
                                        cmdTrain.CommandText = "update days set trainNo='" & txtDTrainNo.Text & "',sun=" & chkSun1.Value & ",mon=" & chkMon1.Value & ",tue=" & chkTue1.Value & ",wed=" & chkWed1.Value & ",thru=" & chkThr1.Value & ",fri=" & chkFri1.Value & ",sat=" & chkSat1.Value & " where trainno='" & frmTrainDialog.downTrainNo & "'"
                                        cmdTrain.Execute
                                        cmdTrain.CommandText = "update days set trainNo='" & txtUTrainNo.Text & "',sun=" & chkSun.Value & ",mon=" & chkMon.Value & ",tue=" & chkTue.Value & ",wed=" & chkWed.Value & ",thru=" & chkThr.Value & ",fri=" & chkFri.Value & ",sat=" & chkSat.Value & " where trainno='" & frmTrainDialog.upTrainNo & "'"
                                        cmdTrain.Execute
                                        
                                        cmdTrain.CommandText = "delete from trainroute where trainno='" & frmTrainDialog.upTrainNo & "' or trainno='" & frmTrainDialog.downTrainNo & "'"
                                        cmdTrain.Execute
                                        flag = 0
                                        For i = 1 To flexGridRoute.Rows - 1
                                            If flexGridRoute.TextMatrix(i, 1) = strChecked Then
                                                flag = flag + 1
                                                cmdTrain.CommandText = "insert into trainroute values('" & txtUTrainNo & "'," & flexGridRoute.TextMatrix(i, 2) & ",'" & flexGridRoute.TextMatrix(i, 4) & "','" & flexGridRoute.TextMatrix(i, 5) & "'," & flag & "," & flexGridRoute.TextMatrix(i, 6) & ")"
                                                cmdTrain.Execute
                                            End If
                                        Next
                                        flag = 0
                                        For i = 1 To flexGridRoute1.Rows - 1
                                            If flexGridRoute1.TextMatrix(i, 1) = strChecked Then
                                                flag = flag + 1
                                                cmdTrain.CommandText = "insert into trainroute values('" & txtDTrainNo & "'," & flexGridRoute1.TextMatrix(i, 2) & ",'" & flexGridRoute1.TextMatrix(i, 4) & "','" & flexGridRoute1.TextMatrix(i, 5) & "'," & flag & "," & flexGridRoute1.TextMatrix(i, 6) & ")"
                                                cmdTrain.Execute
                                            End If
                                        Next
'                                    End If
                                    saveUpdate = 0
                                    MsgBox "Record Successfully Updated", vbInformation
                                End If
                                    txtUTrainNo.Enabled = False
                                    txtDTrainNo.Enabled = False
                                    txtTrainName.Enabled = False
                                    cmbTrainType.Enabled = False
                                    cmbRoute.Enabled = False
                                    chkMon.Enabled = False
                                    chkTue.Enabled = False
                                    chkWed.Enabled = False
                                    chkThr.Enabled = False
                                    chkFri.Enabled = False
                                    chkSat.Enabled = False
                                    chkSun.Enabled = False
                                    chkMon.Value = 0
                                    chkTue.Value = 0
                                    chkWed.Value = 0
                                    chkThr.Value = 0
                                    chkFri.Value = 0
                                    chkSat.Value = 0
                                    chkSun.Value = 0
                                    chkMon1.Enabled = False
                                    chkTue1.Enabled = False
                                    chkWed1.Enabled = False
                                    chkThr1.Enabled = False
                                    chkFri1.Enabled = False
                                    chkSat1.Enabled = False
                                    chkSun1.Enabled = False
                                    chkMon1.Value = 0
                                    chkTue1.Value = 0
                                    chkWed1.Value = 0
                                    chkThr1.Value = 0
                                    chkFri1.Value = 0
                                    chkSat1.Value = 0
                                    chkSun1.Value = 0
                                    txtUTrainNo.BackColor = vbButtonFace
                                    txtDTrainNo.BackColor = vbButtonFace
                                    txtTrainName.BackColor = vbButtonFace
                                    cmbTrainType.BackColor = vbButtonFace
                                    cmbRoute.BackColor = vbButtonFace
                                    txtUTrainNo.Text = ""
                                    txtDTrainNo.Text = ""
                                    txtTrainName.Text = ""
                                    cmbTrainType.ListIndex = -1
                                    cmbRoute.ListIndex = -1
                                    flexGridRoute.Rows = 1
                                    flexGridRoute.Cols = 7
                                    MaskEdBox1.Visible = False
                                    Text1.Visible = False
                                    flexGridRoute1.Rows = 1
                                    flexGridRoute1.Cols = 7
                                    MaskEdBox2.Visible = False
                                    Text2.Visible = False
                            Else
                                MsgBox "Time and Days Sequencing Wrong", vbCritical
                            End If
                        Else
                            MsgBox "Train Must Run Only For Twenty-Four Hours", vbCritical
                        End If
                    Else
                        MsgBox "Train Must Start From Day one", vbCritical
                    End If
                Else
                    MsgBox "Select Atleast Two Station in Route", vbCritical
                End If
            Else
                MsgBox "Atleast One Day Train Must be Run", vbCritical
            End If
        Else
            MsgBox "Please Fill all Fields", vbCritical
        End If
    Else
        MsgBox "Length of Train No. Must Have 5 Digits", vbCritical
    End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub

Private Sub cmbRoute_Click()
    If cmbRoute.ListIndex <> -1 Then
        Set rstRoute = New ADODB.Recordset
        rstRoute.CursorLocation = adUseClient
        rstRoute.Open "select * from routeStn,station where routeID=" & cmbRoute.ItemData(cmbRoute.ListIndex) & " and routestn.stnID=station.stnID order by routestnno asc", railCn
        If rstRoute.RecordCount > 0 Then
            flexGridRoute.Rows = 1
            flexGridRoute.Cols = 7
            flexGridRoute1.Rows = 1
            flexGridRoute1.Cols = 7
            rstRoute.MoveFirst
            For i = 1 To rstRoute.RecordCount
                flexGridRoute.Rows = flexGridRoute.Rows + 1
                flexGridRoute.TextMatrix(i, 0) = flexGridRoute.Rows - 1
                flexGridRoute.TextMatrix(i, 2) = rstRoute("station.stnID")
                flexGridRoute.TextMatrix(i, 3) = rstRoute("stnName")
                flexGridRoute.TextMatrix(i, 4) = "  :  "
                flexGridRoute.TextMatrix(i, 5) = "  :  "
                flexGridRoute.Row = i
                flexGridRoute.Col = 1
                flexGridRoute.CellFontName = "Wingdings"
                flexGridRoute.CellFontSize = 14
                flexGridRoute.CellAlignment = flexAlignCenterCenter
                flexGridRoute.Text = strUnChecked
                rstRoute.MoveNext
            Next
            rstRoute.MoveLast
            For i = 1 To rstRoute.RecordCount
                flexGridRoute1.Rows = flexGridRoute1.Rows + 1
                flexGridRoute1.TextMatrix(i, 0) = flexGridRoute1.Rows - 1
                flexGridRoute1.TextMatrix(i, 2) = rstRoute("station.stnID")
                flexGridRoute1.TextMatrix(i, 3) = rstRoute("stnName")
                flexGridRoute1.TextMatrix(i, 4) = "  :  "
                flexGridRoute1.TextMatrix(i, 5) = "  :  "
                flexGridRoute1.Row = i
                flexGridRoute1.Col = 1
                flexGridRoute1.CellFontName = "Wingdings"
                flexGridRoute1.CellFontSize = 14
                flexGridRoute1.CellAlignment = flexAlignCenterCenter
                flexGridRoute1.Text = strUnChecked
                rstRoute.MovePrevious
            Next
        End If
    End If
End Sub

Private Sub TriggerCheckbox(iRow As Integer, iCol As Integer)
        With flexGridRoute
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
            Else
                .TextMatrix(iRow, iCol) = strUnChecked
            End If
        End With
End Sub
Private Sub TriggerCheckbox1(iRow As Integer, iCol As Integer)
        With flexGridRoute1
            If .TextMatrix(iRow, iCol) = strUnChecked Then
                .TextMatrix(iRow, iCol) = strChecked
            Else
                .TextMatrix(iRow, iCol) = strUnChecked
            End If
        End With
End Sub



Private Sub flexGridRoute_KeyPress(KeyAscii As Integer)
    If flexGridRoute.Rows <> 1 Then
        If flexGridRoute.Col = 1 Then
            If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
                With flexGridRoute
                    Call TriggerCheckbox(.Row, 1)
                End With
            End If
        End If
    End If
End Sub

Private Sub flexGridRoute1_KeyPress(KeyAscii As Integer)
    If flexGridRoute1.Rows <> 1 Then
        If flexGridRoute1.Col = 1 Then
            If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
                With flexGridRoute1
                    Call TriggerCheckbox1(.Row, 1)
                End With
            End If
        End If
    End If
End Sub
Private Sub flexGridRoute_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If flexGridRoute.Rows <> 1 Then
        If flexGridRoute.Col = 1 Then
            If Button = 1 Then
                With flexGridRoute
                    If .MouseRow <> 0 And .MouseCol <> 0 Then
                        Call TriggerCheckbox(.MouseRow, 1)
                    End If
                End With
            End If
        End If
    End If
End Sub
Private Sub flexGridRoute1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If flexGridRoute1.Rows <> 1 Then
        If flexGridRoute1.Col = 1 Then
            If Button = 1 Then
                With flexGridRoute1
                    If .MouseRow <> 0 And .MouseCol <> 0 Then
                        Call TriggerCheckbox1(.MouseRow, 1)
                    End If
                End With
            End If
        End If
    End If
End Sub
Private Sub cmdSearch_Click()
    frmTrainDialog.Show 1
End Sub


Private Sub flexGridRoute1_dblClick()
    If flexGridRoute1.TextMatrix(flexGridRoute1.Row, 1) = strChecked Then
        If flexGridRoute1.Col = 4 Or flexGridRoute1.Col = 5 Then
            GridEdit2 Asc(" ")
        End If
        If flexGridRoute1.Col = 6 Then
            GridEdit3 Asc(" ")
        End If
    End If
End Sub

Private Sub Form_load()
    Set rstTrainType = New ADODB.Recordset
    rstTrainType.CursorLocation = adUseClient
    rstTrainType.Open "select * from traintype", railCn
    If rstTrainType.RecordCount > 0 Then
        i = 0
        rstTrainType.MoveFirst
        Do While Not rstTrainType.EOF
            cmbTrainType.AddItem rstTrainType(1)
            cmbTrainType.ItemData(i) = rstTrainType(0)
            rstTrainType.MoveNext
            i = i + 1
        Loop
    End If
    rstTrainType.Close
    
    Set rstRoute = New ADODB.Recordset
    rstRoute.CursorLocation = adUseClient
    rstRoute.Open "select * from route", railCn
    If rstRoute.RecordCount > 0 Then
        i = 0
        rstRoute.MoveFirst
        Do While Not rstRoute.EOF
            cmbRoute.AddItem rstRoute(1)
            cmbRoute.ItemData(i) = rstRoute(0)
            rstRoute.MoveNext
            i = i + 1
        Loop
    End If
    rstRoute.Close
    Label9.Caption = " Please click Add New Button to Add New Record (OR) Search and Select the Record for Updating Existing Record (OR)"
    flexGridRoute.Rows = 1
    flexGridRoute.Cols = 7
    'flexGridRoute.FixedCols = 1
    flexGridRoute.TextMatrix(0, 0) = "S.No."
    flexGridRoute.TextMatrix(0, 1) = "Select"
    flexGridRoute.TextMatrix(0, 2) = "StnID"
    flexGridRoute.TextMatrix(0, 3) = "Station Name"
    flexGridRoute.TextMatrix(0, 4) = "Arrival"
    flexGridRoute.TextMatrix(0, 5) = "Departure"
    flexGridRoute.TextMatrix(0, 6) = "Day"
    flexGridRoute.ColWidth(0) = 600
    flexGridRoute.ColWidth(1) = 1000
    flexGridRoute.ColWidth(2) = 1000
    flexGridRoute.ColWidth(3) = 2000
    flexGridRoute.ColWidth(4) = 1150
    flexGridRoute.ColWidth(5) = 1150
    flexGridRoute.ColWidth(6) = 500
    txtUTrainNo.Enabled = False
    txtTrainName.Enabled = False
    cmbTrainType.Enabled = False
    cmbRoute.Enabled = False
    chkMon.Enabled = False
    chkTue.Enabled = False
    chkWed.Enabled = False
    chkThr.Enabled = False
    chkFri.Enabled = False
    chkSat.Enabled = False
    chkSun.Enabled = False
    flexGridRoute1.Rows = 1
    flexGridRoute1.Cols = 7
    'flexGridRoute.FixedCols = 1
    flexGridRoute1.TextMatrix(0, 0) = "S.No."
    flexGridRoute1.TextMatrix(0, 1) = "Select"
    flexGridRoute1.TextMatrix(0, 2) = "StnID"
    flexGridRoute1.TextMatrix(0, 3) = "Station Name"
    flexGridRoute1.TextMatrix(0, 4) = "Arrival"
    flexGridRoute1.TextMatrix(0, 5) = "Departure"
    flexGridRoute1.TextMatrix(0, 6) = "Day"
    flexGridRoute1.ColWidth(0) = 600
    flexGridRoute1.ColWidth(1) = 1000
    flexGridRoute1.ColWidth(2) = 1000
    flexGridRoute1.ColWidth(3) = 2000
    flexGridRoute1.ColWidth(4) = 1150
    flexGridRoute1.ColWidth(5) = 1150
    flexGridRoute1.ColWidth(6) = 500
    txtDTrainNo.Enabled = False
    txtTrainName.Enabled = False
    cmbTrainType.Enabled = False
    cmbRoute.Enabled = False
    chkMon1.Enabled = False
    chkTue1.Enabled = False
    chkWed1.Enabled = False
    chkThr1.Enabled = False
    chkFri1.Enabled = False
    chkSat1.Enabled = False
    chkSun1.Enabled = False
    MaskEdBox1.FontName = flexGridRoute.FontName
    MaskEdBox1.FontSize = flexGridRoute.FontSize
    MaskEdBox1.Visible = False
    Text1.FontName = flexGridRoute.FontName
    Text1.FontSize = flexGridRoute.FontSize
    Text1.Visible = False
    MaskEdBox2.FontName = flexGridRoute1.FontName
    MaskEdBox2.FontSize = flexGridRoute1.FontSize
    MaskEdBox2.Visible = False
    Text2.FontName = flexGridRoute1.FontName
    Text2.FontSize = flexGridRoute1.FontSize
    Text2.Visible = False
End Sub

Private Sub GridEdit(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    MaskEdBox1.Move flexGridRoute.CellLeft + flexGridRoute.Left, flexGridRoute.CellTop + flexGridRoute.Top, flexGridRoute.CellWidth, flexGridRoute.CellHeight
    MaskEdBox1.Visible = True
    MaskEdBox1.SetFocus
    
   ' MaskEdBox1.Text = flexGridRoute.TextMatrix()
    Select Case KeyAscii
        Case 0 To Asc(" ")

            MaskEdBox1.SelStart = Len(MaskEdBox1.Text)
            MaskEdBox1.Mask = ""
            MaskEdBox1.Text = ""
            MaskEdBox1.Mask = "##:##"
            If flexGridRoute.Text <> "  :  " Then
                MaskEdBox1.Text = flexGridRoute.Text
            End If
        Case Else
            MaskEdBox1.Text = Chr$(KeyAscii)
            MaskEdBox1.SelStart = 1
    End Select

End Sub
Private Sub GridEdit2(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    MaskEdBox2.Move flexGridRoute1.CellLeft + flexGridRoute1.Left, flexGridRoute1.CellTop + flexGridRoute1.Top, flexGridRoute1.CellWidth, flexGridRoute1.CellHeight
    MaskEdBox2.Visible = True
    MaskEdBox2.SetFocus
    
   ' MaskEdBox1.Text = flexGridRoute.TextMatrix()
    Select Case KeyAscii
        Case 0 To Asc(" ")

            MaskEdBox2.SelStart = Len(MaskEdBox2.Text)
            MaskEdBox2.Mask = ""
            MaskEdBox2.Text = ""
            MaskEdBox2.Mask = "##:##"
            If flexGridRoute1.Text <> "  :  " Then
                MaskEdBox2.Text = flexGridRoute1.Text
            End If
        Case Else
            MaskEdBox2.Text = Chr$(KeyAscii)
            MaskEdBox2.SelStart = 1
    End Select

End Sub





'Private Sub Form_Resize()
'    flexGridRoute.Move 0, 0, ScaleWidth, ScaleHeight
'End Sub


Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            MaskEdBox1.Visible = False
            flexGridRoute.SetFocus

        Case vbKeyReturn
            ' Finish editing.
            flexGridRoute.SetFocus

        Case vbKeyDown
            ' Move down 1 row.
            flexGridRoute.SetFocus
            DoEvents
            If flexGridRoute.Row < flexGridRoute.Rows - 1 Then
                flexGridRoute.Row = flexGridRoute.Row + 1
            End If

        Case vbKeyUp
            ' Move up 1 row.
            flexGridRoute.SetFocus
            DoEvents
            If flexGridRoute.Row > flexGridRoute.FixedRows Then
                flexGridRoute.Row = flexGridRoute.Row - 1
            End If

    End Select
End Sub

Private Sub MaskEdBox2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            MaskEdBox2.Visible = False
            flexGridRoute1.SetFocus

        Case vbKeyReturn
            ' Finish editing.
            flexGridRoute1.SetFocus

        Case vbKeyDown
            ' Move down 1 row.
            flexGridRoute1.SetFocus
            DoEvents
            If flexGridRoute1.Row < flexGridRoute1.Rows - 1 Then
                flexGridRoute1.Row = flexGridRoute1.Row + 1
            End If

        Case vbKeyUp
            ' Move up 1 row.
            flexGridRoute1.SetFocus
            DoEvents
            If flexGridRoute1.Row > flexGridRoute1.FixedRows Then
                flexGridRoute1.Row = flexGridRoute1.Row - 1
            End If

    End Select
End Sub
' Do not beep on Return or Escape.
Private Sub MaskEdBox1_KeyPress(KeyAscii As Integer)

    If Len(MaskEdBox1.ClipText) < 4 Then
        If KeyAscii = vbKeyReturn Then
            If Len(MaskEdBox1.ClipText) <> 4 Then
                MaskEdBox1.Mask = "##:##"
                flexGridRoute.Text = "  :  "
                Exit Sub
            End If
        End If
    End If

    
    Select Case Len(MaskEdBox1.ClipText)

    Case 0
        If Chr(KeyAscii) > 2 Then
            KeyAscii = 0
        End If
    Case 2
        If KeyAscii = 8 Then
            Exit Sub
        ElseIf Chr(KeyAscii) > 5 Then
            KeyAscii = 0
        End If
    Case 1
        If KeyAscii = 8 Then
            Exit Sub
        ElseIf Chr(KeyAscii) > 3 Then
            If Mid(MaskEdBox1.ClipText, 1, 1) = 2 Then
                KeyAscii = 0
                
            End If
        End If

    End Select
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub
Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)

    If Len(MaskEdBox2.ClipText) < 4 Then
        If KeyAscii = vbKeyReturn Then
            If Len(MaskEdBox2.ClipText) <> 4 Then
                MaskEdBox2.Mask = "##:##"
                flexGridRoute1.Text = "  :  "
                Exit Sub
            End If
        End If
    End If

    
    Select Case Len(MaskEdBox2.ClipText)

    Case 0
        If Chr(KeyAscii) > 2 Then
            KeyAscii = 0
        End If
    Case 2
        If KeyAscii = 8 Then
            Exit Sub
        ElseIf Chr(KeyAscii) > 5 Then
            KeyAscii = 0
        End If
    Case 1
        If KeyAscii = 8 Then
            Exit Sub
        ElseIf Chr(KeyAscii) > 3 Then
            If Mid(MaskEdBox2.ClipText, 1, 1) = 2 Then
                KeyAscii = 0
                
            End If
        End If

    End Select
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

Private Sub flexGridRoute_DblClick()
    If flexGridRoute.TextMatrix(flexGridRoute.Row, 1) = strChecked Then
        If flexGridRoute.Col = 4 Or flexGridRoute.Col = 5 Then
            GridEdit Asc(" ")
        End If
        If flexGridRoute.Col = 6 Then
            GridEdit1 Asc(" ")
        End If
    End If
End Sub

'Private Sub flexGridRoute_KeyPress(KeyAscii As Integer)
'    GridEdit KeyAscii
'End Sub

Private Sub flexGridRoute_LeaveCell()
    If MaskEdBox1.Visible Then
        flexGridRoute.Text = MaskEdBox1.Text
        MaskEdBox1.Visible = False
        If Len(MaskEdBox1.ClipText) <> 4 Then
            MaskEdBox1.Mask = "##:##"
            flexGridRoute.Text = "  :  "
        End If
    End If
    If Text1.Visible Then
        flexGridRoute.Text = Text1.Text
        Text1.Visible = False
    End If
End Sub
Private Sub flexGridRoute1_LeaveCell()
    If MaskEdBox2.Visible Then
        flexGridRoute1.Text = MaskEdBox2.Text
        MaskEdBox2.Visible = False
        If Len(MaskEdBox2.ClipText) <> 4 Then
            MaskEdBox2.Mask = "##:##"
            flexGridRoute1.Text = "  :  "
        End If
    End If
    If Text2.Visible Then
        flexGridRoute1.Text = Text2.Text
        Text2.Visible = False
    End If
End Sub

Private Function checkEmptyFields() As Boolean
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 1) = strChecked Then
            If Mid(flexGridRoute.TextMatrix(i, 4), 1, 1) = "" Or Mid(flexGridRoute.TextMatrix(i, 5), 1, 1) = "" Or flexGridRoute.TextMatrix(i, 6) = "" Then
                checkEmptyFields = False
                Exit Function
            End If
        End If
        If flexGridRoute1.TextMatrix(i, 1) = strChecked Then
            If Mid(flexGridRoute1.TextMatrix(i, 4), 1, 1) = "" Or Mid(flexGridRoute1.TextMatrix(i, 5), 1, 1) = "" Or flexGridRoute1.TextMatrix(i, 6) = "" Then
                checkEmptyFields = False
                Exit Function
            End If
        End If
    Next
    checkEmptyFields = True
End Function
Private Sub flexGridRoute_GotFocus()
    If flexGridRoute.Rows <> 1 Then
        If MaskEdBox1.Visible Then
            If Len(MaskEdBox1.ClipText) <> 4 Then
                
                MaskEdBox1.Mask = ""
                MaskEdBox1.Text = ""
                MaskEdBox1.Mask = "##:##"
                flexGridRoute.Text = "  :  "
                MaskEdBox1.Visible = False
            Else
                flexGridRoute.Text = MaskEdBox1.Text
                MaskEdBox1.Visible = False
            End If
            'if flexgridroute.TextMatrix( )
        End If
        If Text1.Visible Then
            flexGridRoute.Text = Text1.Text
            Text1.Visible = False
        End If
    End If
End Sub

Private Sub flexGridRoute1_GotFocus()
    If flexGridRoute1.Rows <> 1 Then
        If MaskEdBox2.Visible Then
            If Len(MaskEdBox2.ClipText) <> 4 Then
                
                MaskEdBox2.Mask = ""
                MaskEdBox2.Text = ""
                MaskEdBox2.Mask = "##:##"
                flexGridRoute1.Text = "  :  "
                MaskEdBox2.Visible = False
            Else
                flexGridRoute1.Text = MaskEdBox2.Text
                MaskEdBox2.Visible = False
            End If
            'if flexgridroute.TextMatrix( )
        End If
        If Text2.Visible Then
            flexGridRoute1.Text = Text2.Text
            Text2.Visible = False
        End If
    End If
End Sub
Private Sub GridEdit1(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    Text1.Left = flexGridRoute.CellLeft + flexGridRoute.Left
    Text1.Top = flexGridRoute.CellTop + flexGridRoute.Top
    Text1.Width = flexGridRoute.CellWidth
    Text1.Height = flexGridRoute.CellHeight
    Text1.Visible = True
    Text1.SetFocus

    Select Case KeyAscii
        Case 0 To Asc(" ")
            Text1.Text = flexGridRoute.Text
            Text1.SelStart = Len(Text1.Text)
        Case Else
            Text1.Text = Chr$(KeyAscii)
            Text1.SelStart = 1
    End Select

End Sub
Private Sub GridEdit3(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    Text2.Left = flexGridRoute1.CellLeft + flexGridRoute1.Left
    Text2.Top = flexGridRoute1.CellTop + flexGridRoute1.Top
    Text2.Width = flexGridRoute1.CellWidth
    Text2.Height = flexGridRoute1.CellHeight
    Text2.Visible = True
    Text2.SetFocus

    Select Case KeyAscii
        Case 0 To Asc(" ")
            Text2.Text = flexGridRoute1.Text
            Text2.SelStart = Len(Text2.Text)
        Case Else
            Text2.Text = Chr$(KeyAscii)
            Text2.SelStart = 1
    End Select

End Sub
'Private Sub Form_Resize()
'    flexgridroute.Move 0, 0, ScaleWidth, ScaleHeight
'End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            Text1.Visible = False
            flexGridRoute.SetFocus

        Case vbKeyReturn
            ' Finish editing.
            flexGridRoute.SetFocus

        Case vbKeyDown
            ' Move down 1 row.
            flexGridRoute.SetFocus
            DoEvents
            If flexGridRoute.Row < flexGridRoute.Rows - 1 Then
                flexGridRoute.Row = flexGridRoute.Row + 1
            End If

        Case vbKeyUp
            ' Move up 1 row.
            flexGridRoute.SetFocus
            DoEvents
            If flexGridRoute.Row > flexGridRoute.FixedRows Then
                flexGridRoute.Row = flexGridRoute.Row - 1
            End If

    End Select
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            Text2.Visible = False
            flexGridRoute1.SetFocus

        Case vbKeyReturn
            ' Finish editing.
            flexGridRoute1.SetFocus

        Case vbKeyDown
            ' Move down 1 row.
            flexGridRoute1.SetFocus
            DoEvents
            If flexGridRoute1.Row < flexGridRoute1.Rows - 1 Then
                flexGridRoute1.Row = flexGridRoute1.Row + 1
            End If

        Case vbKeyUp
            ' Move up 1 row.
            flexGridRoute1.SetFocus
            DoEvents
            If flexGridRoute1.Row > flexGridRoute1.FixedRows Then
                flexGridRoute1.Row = flexGridRoute1.Row - 1
            End If

    End Select
End Sub
' Do not beep on Return or Escape.
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
        Call validation(1, KeyAscii, Text1)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
        Call validation(1, KeyAscii, Text1)
End Sub


Private Sub Timer2_Timer()
    strs = Mid(Label9.Caption, 1, 1)
    Label9.Caption = Mid(Label9.Caption, 2, Len(Label9.Caption)) & strs
    If Label7.ForeColor = vbYellow Then
        Label7.ForeColor = vbHighlight
    ElseIf Label7.ForeColor = &H8080FF Then
        Label7.ForeColor = vbYellow
    ElseIf Label7.ForeColor = vbHighlight Then
        Label7.ForeColor = vbGreen
    Else
        Label7.ForeColor = &H8080FF
    End If
End Sub

Private Sub txtTrainName_KeyPress(KeyAscii As Integer)
    Call validation(2, KeyAscii, txtTrainName)
End Sub

Private Sub txtdTrainNo_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(txtDTrainNo.Text) < 5 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            txtDTrainNo.Text = txtDTrainNo.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        txtDTrainNo.Text = txtDTrainNo.Text & Chr(KeyAscii)
    End If
End If
End Sub

Private Sub txtTrainName_LostFocus()
   If txtTrainName <> "" Then
       Set rstTrain = New ADODB.Recordset
        rstTrain.CursorLocation = adUseClient
        rstTrain.Open "select trainname from train", railCn
        If rstTrain.RecordCount > 0 Then
            i = 0
            rstTrain.MoveFirst
            Do While Not rstTrain.EOF
                If saveUpdate = 2 Then
                    If txtTrainName.Text = rstTrain("trainname") And rstTrain("trainname") <> frmTrainDialog.trainName Then
                        MsgBox "Train Name Already Exits, Please Give Another Train Name..", vbCritical
                        txtTrainName.Text = ""
                        txtTrainName.SetFocus
                        Exit Sub
                    Else
                        rstTrain.MoveNext
                    End If
                ElseIf saveUpdate = 1 Then
                    If txtTrainName.Text = rstTrain("trainname") Then
                        MsgBox "Train Name Already Exits, Please Give Another Train Name.", vbCritical
                        txtTrainName.Text = ""
                        txtTrainName.SetFocus
                        Exit Sub
                    Else
                        rstTrain.MoveNext
                    End If
    
                End If
            Loop
        End If
    End If
End Sub

Private Sub txtuTrainNo_LostFocus()
    If txtUTrainNo <> "" Then
        Set rstTrain = New ADODB.Recordset
        rstTrain.CursorLocation = adUseClient
        rstTrain.Open "select downtrainno,uptrainno from train", railCn
        If rstTrain.RecordCount > 0 Then
            i = 0
            rstTrain.MoveFirst
            Do While Not rstTrain.EOF
                If saveUpdate = 2 Then
                    If (txtUTrainNo.Text = rstTrain("uptrainno") Or txtUTrainNo.Text = rstTrain("downtrainno")) And (rstTrain("uptrainno") <> frmTrainDialog.upTrainNo Or txtUTrainNo.Text = frmTrainDialog.downTrainNo) Then
                        MsgBox "Train No. Already Exits, Please Give Another Train No..", vbCritical
                        txtUTrainNo.Text = ""
                        txtUTrainNo.SetFocus
                        Exit Sub
                    Else
                        rstTrain.MoveNext
                    End If
                ElseIf saveUpdate = 1 Then
                    If txtUTrainNo.Text = rstTrain("uptrainno") Then
                        MsgBox "Train No. Already Exits, Please Give Another Train No..", vbCritical
                        txtUTrainNo.Text = ""
                        txtUTrainNo.SetFocus
                        Exit Sub
                    Else
                        rstTrain.MoveNext
                    End If
    
                End If
            Loop
        End If
    End If
End Sub
Private Sub txtdTrainNo_LostFocus()
   If txtDTrainNo <> "" Then
       Set rstTrain = New ADODB.Recordset
        rstTrain.CursorLocation = adUseClient
        rstTrain.Open "select uptrainno,downtrainno from train", railCn
        If rstTrain.RecordCount > 0 Then
            i = 0
            rstTrain.MoveFirst
            Do While Not rstTrain.EOF
                If saveUpdate = 2 Then
                    If (txtDTrainNo.Text = rstTrain("downtrainno") Or txtDTrainNo.Text = rstTrain("uptrainno")) And (rstTrain("downtrainno") <> frmTrainDialog.downTrainNo Or txtDTrainNo.Text = frmTrainDialog.upTrainNo) Then
                        MsgBox "Train No. Already Exits, Please Give Another Train No..", vbCritical
                        txtDTrainNo.Text = ""
                        txtDTrainNo.SetFocus
                        Exit Sub
                    Else
                        rstTrain.MoveNext
                    End If
                ElseIf saveUpdate = 1 Then
                    If txtDTrainNo.Text = rstTrain("downtrainno") Then
                        MsgBox "Train No. Already Exits, Please Give Another Train No..", vbCritical
                        txtDTrainNo.Text = ""
                        txtDTrainNo.SetFocus
                        Exit Sub
                    Else
                        rstTrain.MoveNext
                    End If
    
                End If
            Loop
        End If
    End If
End Sub
Private Sub txtUTrainNo_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(txtUTrainNo.Text) < 5 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            txtUTrainNo.Text = txtUTrainNo.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        txtUTrainNo.Text = txtUTrainNo.Text & Chr(KeyAscii)
    End If
End If
End Sub
