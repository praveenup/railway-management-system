VERSION 5.00
Begin VB.Form frmCoachType 
   BackColor       =   &H8000000E&
   Caption         =   "Coach Type"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
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
      Height          =   6375
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   16095
      Begin VB.TextBox txtCoachType 
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
         Left            =   2640
         TabIndex        =   22
         Top             =   600
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         Caption         =   "Add Seat Berth Type(Cabin)"
         BeginProperty Font 
            Name            =   "Forte"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   5655
         Left            =   7320
         TabIndex        =   20
         Top             =   360
         Width           =   8175
         Begin VB.ComboBox cmbBox 
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
            Index           =   11
            Left            =   6000
            TabIndex        =   45
            Top             =   4800
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   10
            Left            =   1920
            TabIndex        =   43
            Top             =   4800
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   9
            Left            =   6000
            TabIndex        =   41
            Top             =   3960
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   8
            Left            =   1920
            TabIndex        =   39
            Top             =   3960
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   7
            Left            =   6000
            TabIndex        =   37
            Top             =   3120
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   6
            Left            =   1920
            TabIndex        =   35
            Top             =   3120
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   5
            Left            =   6000
            TabIndex        =   33
            Top             =   2280
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   4
            Left            =   1920
            TabIndex        =   31
            Top             =   2280
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   3
            Left            =   6000
            TabIndex        =   29
            Top             =   1440
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   2
            Left            =   1920
            TabIndex        =   27
            Top             =   1440
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   1
            Left            =   6000
            TabIndex        =   25
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cmbBox 
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
            Index           =   0
            Left            =   1920
            TabIndex        =   23
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Height          =   375
            Left            =   14280
            Picture         =   "frmCoachType.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 12:"
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
            Index           =   11
            Left            =   4560
            TabIndex        =   46
            Top             =   4800
            Width           =   1455
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 11 :"
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
            Index           =   10
            Left            =   600
            TabIndex        =   44
            Top             =   4800
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 10 :"
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
            Index           =   9
            Left            =   4560
            TabIndex        =   42
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 9:"
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
            Index           =   8
            Left            =   600
            TabIndex        =   40
            Top             =   3960
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 8 :"
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
            Index           =   7
            Left            =   4560
            TabIndex        =   38
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 7 :"
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
            Index           =   6
            Left            =   600
            TabIndex        =   36
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 6 :"
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
            Index           =   5
            Left            =   4560
            TabIndex        =   34
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 5 :"
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
            Index           =   4
            Left            =   600
            TabIndex        =   32
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 4 :"
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
            Index           =   3
            Left            =   4560
            TabIndex        =   30
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 3 :"
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
            Index           =   2
            Left            =   600
            TabIndex        =   28
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 2 :"
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
            Index           =   1
            Left            =   4560
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblBerth 
            BackColor       =   &H8000000E&
            Caption         =   "Seat No. 1 :"
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
            Index           =   0
            Left            =   600
            TabIndex        =   24
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkPassenger 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtCabins 
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
         Left            =   2640
         TabIndex        =   18
         Top             =   4680
         Width           =   855
      End
      Begin VB.ComboBox cmbCoachInitial 
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
         Left            =   2640
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkChair 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox chkAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   12
         Top             =   3240
         Width           =   735
      End
      Begin VB.ComboBox cmbCabinSeat 
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
         Left            =   2640
         TabIndex        =   0
         Top             =   5640
         Width           =   855
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   5640
         Picture         =   "frmCoachType.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000E&
         Caption         =   "Coach Initial:"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000E&
         Caption         =   "Passenger coach:"
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
         TabIndex        =   13
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Chair Car:"
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
         TabIndex        =   11
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Number of Seats in per Cabins:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Number of Cabins:"
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
         TabIndex        =   9
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "AC Availability:"
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
         TabIndex        =   8
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Coach Type Name:"
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
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   4080
      Picture         =   "frmCoachType.frx":20B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   9480
      Picture         =   "frmCoachType.frx":47F2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   5880
      Picture         =   "frmCoachType.frx":70F0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   7680
      Picture         =   "frmCoachType.frx":996A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   1800
      Picture         =   "frmCoachType.frx":C268
      Top             =   480
      Width           =   5985
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   1320
      Picture         =   "frmCoachType.frx":1C47A
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label21 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD COACH TYPE INFORMATION"
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
      Left            =   2040
      TabIndex        =   17
      Top             =   0
      Width           =   6495
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmCoachType.frx":1CCD6
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11520
      Picture         =   "frmCoachType.frx":1D1B0
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmCoachType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdCoach As ADODB.Command
Dim rstCoach As ADODB.Recordset
Dim rstCoach1 As ADODB.Recordset
Dim rstBerthType As ADODB.Recordset
Dim ID As Integer
Dim flag As Integer   'see insert at save
Dim strBerth(12) As Integer 'for storing berthtype ,see insert statement at save button
Private Sub CalculateSeats()
Dim flag As Integer
Dim strBerth(12) As String * 10
For i = 0 To cmbCabinSeat.ListIndex
    strBerth(i) = cmbBox(i).ItemData(cmbBox(i).ListIndex)
Next
For i = 0 To Val(txtCabins.Text)
    For j = 0 To cmbCabinSeat.ListIndex
        flag = flag + 1
        seatno = flag
        berthtype = strBerth(j)
    Next
Next
End Sub

Private Sub chkpassenger_Click()
    If chkPassenger = 1 Then
        cmdNew.Top = 7920
        cmdSave.Top = 7920
        cmdDelete.Top = 7920
        cmdClose.Top = 7920
        Frame1.Height = 6200
    Else
        Frame1.Height = 3000
        cmdNew.Left = 960
        cmdNew.Top = 4560
        cmdSave.Left = 2760
        cmdSave.Top = 4560
        cmdDelete.Left = 4560
        cmdDelete.Top = 4560
        cmdClose.Left = 6360
        cmdClose.Top = 4560
        Frame1.Width = 7215
        txtCabins.Text = ""
        chkAC.Value = 0
        chkChair.Value = 0
        cmbCabinSeat.ListIndex = -1
    End If
End Sub


Private Sub cmbCabinSeat_Click()
    If cmbCabinSeat.ListIndex <> -1 Then
        Frame1.Width = 16095
        
        For i = 0 To 11
            cmbBox(i).Visible = False
            lblBerth(i).Visible = False
            cmbBox(i).ListIndex = -1
        Next
        For i = 0 To cmbCabinSeat.ItemData(cmbCabinSeat.ListIndex) - 1
            cmbBox(i).Visible = True
            lblBerth(i).Visible = True
        Next
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
    If saveUpdate = 2 Then
        If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
                Set cmdCoach = New ADODB.Command
                cmdCoach.CommandType = adCmdText
                cmdCoach.ActiveConnection = railCn
                cmdCoach.CommandText = "delete from coachBerth where coachTypeID=" & frmCoachTypeDialog.coachID & ""
                cmdCoach.Execute
                cmdCoach.CommandText = "delete from coach where coachTypeID=" & frmCoachTypeDialog.coachID & ""
                cmdCoach.Execute

                MsgBox "Record Successfully Deleted", vbInformation
                Frame1.Height = 3000
                Frame1.Width = 7215
                txtCoachType.Text = ""
                txtCabins.Text = ""
                cmbCoachInitial.ListIndex = -1
                cmbCoachInitial.Text = ""
                cmbCabinSeat.ListIndex = -1
                chkPassenger.Value = 0
                chkAC.Value = 0
                chkChair.Value = 0
                txtCoachType.Enabled = False
                cmbCoachInitial.Enabled = False
                chkPassenger.Enabled = False
                cmdNew.Left = 960
                cmdNew.Top = 4560
                cmdSave.Left = 2760
                cmdSave.Top = 4560
                cmdDelete.Left = 4560
                cmdDelete.Top = 4560
                cmdClose.Left = 6360
                cmdClose.Top = 4560
                txtCoachType.BackColor = vbButtonFace
                cmbCoachInitial.BackColor = vbButtonFace
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
    Frame1.Height = 3000
    Frame1.Width = 7215
    txtCoachType.Text = ""
    cmbCoachInitial.ListIndex = -1
    chkPassenger.Value = 0
    cmbCoachInitial.Text = ""
    txtCoachType.Enabled = True
    cmbCoachInitial.Enabled = True
    chkPassenger.Enabled = True
    txtCoachType.BackColor = vbHighlightText
    cmbCoachInitial.BackColor = vbHighlightText
    cmdNew.Left = 960
    cmdNew.Top = 4560
    cmdSave.Left = 2760
    cmdSave.Top = 4560
    cmdDelete.Left = 4560
    cmdDelete.Top = 4560
    cmdClose.Left = 6360
    cmdClose.Top = 4560
    
    saveUpdate = 1
End Sub

Private Function checkSeatType() As Boolean
    If cmbCabinSeat.ListIndex <> -1 Then
        For i = 0 To cmbCabinSeat.ListIndex
            If cmbBox(i).Text = "" Then
                checkSeatType = False
                Exit For
            Else
                checkSeatType = True
            End If
        Next
    End If
End Function
Private Function checkAlready() As Boolean
    Set rstCoach1 = New ADODB.Recordset
    rstCoach1.CursorLocation = adUseClient
    rstCoach1.Open "select * from coach ", railCn
    If rstCoach1.RecordCount > 0 Then
        rstCoach1.MoveFirst
        For i = 0 To rstCoach1.RecordCount - 1
            If rstCoach1("CoachTypeName") = txtCoachType.Text Or rstCoach1("coachinitial") = cmbCoachInitial.Text Then
                checkAlready = False
                rstCoach1.Close
                Exit Function
            End If
            rstCoach1.MoveNext
        Next
    End If
    checkAlready = True
    rstCoach1.Close
End Function

Private Function checkAlready1() As Boolean
    Set rstCoach1 = New ADODB.Recordset
    rstCoach1.CursorLocation = adUseClient
    rstCoach1.Open "select * from coach ", railCn
    If rstCoach1.RecordCount > 0 Then
        rstCoach1.MoveFirst
        For i = 0 To rstCoach1.RecordCount - 1
            If (rstCoach1("CoachTypeName") = txtCoachType.Text Or rstCoach1("coachinitial") = cmbCoachInitial.Text) And (cmbCoachInitial.Text <> frmCoachTypeDialog.coachTypeInitial Or frmCoachTypeDialog.coachTypeName <> txtCoachType.Text) Then
                checkAlready1 = False
                rstCoach1.Close
                Exit Function
            End If
            rstCoach1.MoveNext
        Next
    End If
    checkAlready1 = True
    rstCoach1.Close
End Function
Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
            Set cmdCoach = New ADODB.Command
            Set rstCoach = New ADODB.Recordset
            cmdCoach.CommandType = adCmdText
            cmdCoach.ActiveConnection = railCn
            If saveUpdate = 1 Then
                If checkAlready() Then
                    rstCoach.Open "select max(coachtypeID) from coach", railCn
                    If rstCoach.Fields(0) > 0 Then
                        ID = rstCoach.Fields(0) + 1
                    Else
                        ID = 1
                    End If
                    rstCoach.Close
                    
                    If chkPassenger.Value = 0 Then
                        If txtCoachType <> "" And cmbCoachInitial.Text <> "" Then
                            cmdCoach.CommandText = "insert into coach (coachtypeid,coachtypename,coachinitial)values(" & ID & ",'" & txtCoachType.Text & "','" & cmbCoachInitial.Text & "')"
                            cmdCoach.Execute
                            saveUpdate = 0
                            MsgBox "Record Successfully Saved", vbInformation
                        Else
                            MsgBox "Please Fill all Fields", vbCritical
                            GoTo label
                        End If
                    Else
                        If txtCoachType <> "" And cmbCoachInitial.Text <> "" And cmbCabinSeat.ListIndex <> -1 And txtCabins.Text <> "" And checkSeatType() Then
                            cmdCoach.CommandText = "insert into coach values(" & ID & ",'" & txtCoachType.Text & "','" & cmbCoachInitial.Text & "'," & True & "," & chkAC.Value & "," & chkChair.Value & "," & Val(txtCabins.Text) & "," & cmbCabinSeat.ItemData(cmbCabinSeat.ListIndex) & ")"
                            cmdCoach.Execute
                            flag = 0
                            For i = 0 To cmbCabinSeat.ListIndex
                                strBerth(i) = cmbBox(i).ItemData(cmbBox(i).ListIndex)
                            Next
                            
                            For i = 0 To Val(txtCabins.Text) - 1
                                For j = 0 To cmbCabinSeat.ListIndex
                                    flag = flag + 1
                                    cmdCoach.CommandText = "insert into coachBerth values(" & ID & "," & flag & "," & strBerth(j) & ")"
                                    cmdCoach.Execute
                                Next
                            Next
                            saveUpdate = 0
                            MsgBox "Record Successfully Saved", vbInformation
                        Else
                            MsgBox "Please Fill all Fields", vbCritical
                            GoTo label
                        End If
                    End If
                    Else
                        MsgBox "Coach Information Already Exists, Please Give Another CoachName or CoachInitial", vbCritical
                        GoTo label
                    End If
                ElseIf saveUpdate = 2 Then
                    If checkAlready1() Then
                        If chkPassenger.Value = 0 Then
                            If txtCoachType <> "" And cmbCoachInitial.Text <> "" Then
                                cmdCoach.CommandText = "update coach set coachtypename='" & txtCoachType.Text & "',coachinitial='" & cmbCoachInitial.Text & "' where coachTypeID=" & frmCoachTypeDialog.coachID & ""
                                cmdCoach.Execute
                                saveUpdate = 0
                                MsgBox "Record Successfully Updated", vbInformation
                            Else
                                MsgBox "Please Fill all Fields", vbCritical
                                GoTo label
                            End If
                        Else
                            If txtCoachType <> "" And cmbCoachInitial.Text <> "" And cmbCabinSeat.ListIndex <> -1 And txtCabins.Text <> "" And checkSeatType() Then
                                cmdCoach.CommandText = "update coach set coachtypename='" & txtCoachType.Text & "',coachinitial='" & cmbCoachInitial.Text & "',passengercoach=" & True & ",ac=" & chkAC.Value & ",chaircar=" & chkChair.Value & ",cabin=" & Val(txtCabins.Text) & ",cabinseats=" & cmbCabinSeat.ItemData(cmbCabinSeat.ListIndex) & " where coachTypeID=" & frmCoachTypeDialog.coachID & ""
                                cmdCoach.Execute
                                cmdCoach.CommandText = "delete from coachBerth where coachTypeID=" & frmCoachTypeDialog.coachID & ""
                                cmdCoach.Execute
                                flag = 0
                                For i = 0 To cmbCabinSeat.ListIndex
                                    strBerth(i) = cmbBox(i).ItemData(cmbBox(i).ListIndex)
                                Next
                                
                                For i = 0 To Val(txtCabins.Text) - 1
                                    For j = 0 To cmbCabinSeat.ListIndex
                                        flag = flag + 1
                                        cmdCoach.CommandText = "insert into coachBerth values(" & frmCoachTypeDialog.coachID & "," & flag & "," & strBerth(j) & ")"
                                        cmdCoach.Execute
                                    Next
                                Next
                                saveUpdate = 0
                                MsgBox "Record Successfully Updated", vbInformation
                            Else
                                MsgBox "Please Fill all Fields", vbCritical
                                GoTo label
                            End If
                        End If
                    Else
                        MsgBox "Coach Information Already Exists, Please Give Another CoachName or CoachInitial", vbCritical
                        GoTo label
                    End If
                End If
                    Frame1.Height = 3000
                    Frame1.Width = 7215
                    txtCoachType.Text = ""
                    txtCabins.Text = ""
                    cmbCoachInitial.ListIndex = -1
                    cmbCabinSeat.ListIndex = -1
                    chkPassenger.Value = 0
                    chkAC.Value = 0
                    chkChair.Value = 0
                    cmbCoachInitial.Text = ""
                    txtCoachType.Enabled = False
                    cmbCoachInitial.Enabled = False
                    chkPassenger.Enabled = False
                    cmdNew.Left = 960
                    cmdNew.Top = 4560
                    cmdSave.Left = 2760
                    cmdSave.Top = 4560
                    cmdDelete.Left = 4560
                    cmdDelete.Top = 4560
                    cmdClose.Left = 6360
                    cmdClose.Top = 4560
                    txtCoachType.BackColor = vbButtonFace
                    cmbCoachInitial.BackColor = vbButtonFace

label:
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub

Private Sub cmdSearch_Click()
    frmCoachTypeDialog.Show 1
End Sub

Private Sub Form_load()
    Frame1.Height = 3000
    Frame1.Width = 7215
    txtCoachType.Text = ""
    cmbCoachInitial.ListIndex = -1
    txtCoachType.Enabled = False
    cmbCoachInitial.Enabled = False
    chkPassenger.Enabled = False
    cmdNew.Left = 960
    cmdNew.Top = 4560
    cmdSave.Left = 2760
    cmdSave.Top = 4560
    cmdDelete.Left = 4560
    cmdDelete.Top = 4560
    cmdClose.Left = 6360
    cmdClose.Top = 4560
    
    For i = 65 To 90
        cmbCoachInitial.AddItem Chr(i)
    Next
    
    For i = 1 To 12
        cmbCabinSeat.AddItem i
        cmbCabinSeat.ItemData(i - 1) = i
    Next
    Set rstBerthType = New ADODB.Recordset
    rstBerthType.CursorLocation = adUseClient
    rstBerthType.Open "select berthID,berthType from berthtype", railCn
    If rstBerthType.RecordCount > 0 Then
        rstBerthType.MoveFirst
        For i = 0 To rstBerthType.RecordCount - 1
            For j = 0 To 11
                cmbBox(j).AddItem rstBerthType(1)
                cmbBox(j).ItemData(i) = rstBerthType(0)
            Next
            rstBerthType.MoveNext
        Next
    End If
End Sub



Private Sub txtCabins_KeyPress(KeyAscii As Integer)
    Call validation(1, KeyAscii, txtCabins)
End Sub

Private Sub txtCoachType_KeyPress(KeyAscii As Integer)
    Call validation(2, KeyAscii, txtCoachType)
End Sub
