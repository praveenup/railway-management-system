VERSION 5.00
Begin VB.Form frmReservation 
   BackColor       =   &H8000000E&
   Caption         =   "Ticket Reservation"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   19020
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdBook 
      Height          =   375
      Left            =   7200
      Picture         =   "frmReservation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7800
      Width           =   3135
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
      Height          =   6855
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   16455
      Begin VB.TextBox txtContact 
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
         Left            =   7800
         TabIndex        =   28
         Top             =   6000
         Width           =   1815
      End
      Begin VB.ComboBox cmbGender 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         ItemData        =   "frmReservation.frx":4356
         Left            =   10680
         List            =   "frmReservation.frx":4358
         TabIndex        =   26
         Top             =   3840
         Width           =   855
      End
      Begin VB.ComboBox cmbGender 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         ItemData        =   "frmReservation.frx":435A
         Left            =   10680
         List            =   "frmReservation.frx":435C
         TabIndex        =   25
         Top             =   4320
         Width           =   855
      End
      Begin VB.ComboBox cmbGender 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         ItemData        =   "frmReservation.frx":435E
         Left            =   10680
         List            =   "frmReservation.frx":4360
         TabIndex        =   24
         Top             =   4800
         Width           =   855
      End
      Begin VB.ComboBox cmbGender 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         ItemData        =   "frmReservation.frx":4362
         Left            =   10680
         List            =   "frmReservation.frx":4364
         TabIndex        =   23
         Top             =   5280
         Width           =   855
      End
      Begin VB.ComboBox cmbGender 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "frmReservation.frx":4366
         Left            =   10680
         List            =   "frmReservation.frx":4368
         TabIndex        =   22
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   9600
         TabIndex        =   21
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   9600
         TabIndex        =   20
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   9600
         TabIndex        =   19
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   9600
         TabIndex        =   18
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   9600
         TabIndex        =   17
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   5880
         TabIndex        =   16
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   5880
         TabIndex        =   15
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5880
         TabIndex        =   14
         Top             =   4800
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   5880
         TabIndex        =   13
         Top             =   5280
         Width           =   3015
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   12
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label17 
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
         TabIndex        =   36
         Top             =   360
         Width           =   15975
      End
      Begin VB.Label Label16 
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
         TabIndex        =   35
         Top             =   2160
         Width           =   15975
      End
      Begin VB.Image Image7 
         Height          =   405
         Left            =   240
         Picture         =   "frmReservation.frx":436A
         Top             =   2160
         Width           =   16005
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   240
         Picture         =   "frmReservation.frx":19598
         Top             =   360
         Width           =   16005
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
         Left            =   4440
         TabIndex        =   34
         Top             =   1560
         Width           =   3615
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
         Left            =   10920
         TabIndex        =   33
         Top             =   1560
         Width           =   3375
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
         TabIndex        =   32
         Top             =   1080
         Width           =   2895
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
         TabIndex        =   31
         Top             =   960
         Width           =   2655
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
         Left            =   1560
         TabIndex        =   30
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Line Line7 
         BorderWidth     =   5
         X1              =   4320
         X2              =   11760
         Y1              =   2760
         Y2              =   2760
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
         TabIndex        =   29
         Top             =   6120
         Width           =   1935
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
         Left            =   4680
         TabIndex        =   11
         Top             =   3840
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
         Left            =   4680
         TabIndex        =   10
         Top             =   4320
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
         Left            =   4680
         TabIndex        =   9
         Top             =   4800
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
         Left            =   4680
         TabIndex        =   8
         Top             =   5280
         Width           =   255
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
         Left            =   4680
         TabIndex        =   7
         Top             =   3360
         Width           =   255
      End
      Begin VB.Line Line6 
         BorderWidth     =   5
         X1              =   4320
         X2              =   11760
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line5 
         BorderWidth     =   5
         X1              =   10440
         X2              =   10440
         Y1              =   2760
         Y2              =   5880
      End
      Begin VB.Line Line4 
         BorderWidth     =   5
         X1              =   9360
         X2              =   9360
         Y1              =   2760
         Y2              =   5880
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   5400
         X2              =   5400
         Y1              =   2760
         Y2              =   5880
      End
      Begin VB.Line Line2 
         BorderWidth     =   5
         X1              =   11760
         X2              =   11760
         Y1              =   2760
         Y2              =   5880
      End
      Begin VB.Line Line1 
         BorderWidth     =   5
         X1              =   4320
         X2              =   4320
         Y1              =   2760
         Y2              =   5880
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C000&
         Caption         =   "    S.No.           Name                                                    Age        Gender   "
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
         Left            =   4320
         TabIndex        =   6
         Top             =   2760
         Width           =   7455
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
         Left            =   9600
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
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
         Left            =   2760
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
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
         Height          =   495
         Left            =   12840
         TabIndex        =   3
         Top             =   960
         Width           =   1935
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
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
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
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Image Image8 
         Height          =   1485
         Left            =   240
         Picture         =   "frmReservation.frx":2E7C6
         Top             =   720
         Width           =   16005
      End
      Begin VB.Image Image9 
         Height          =   4245
         Left            =   240
         Picture         =   "frmReservation.frx":7BF14
         Top             =   2400
         Width           =   16005
      End
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   240
      Picture         =   "frmReservation.frx":159542
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Reservation"
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
      Left            =   1080
      TabIndex        =   37
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmReservation.frx":159E8D
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11520
      Picture         =   "frmReservation.frx":15A367
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   17280
      Picture         =   "frmReservation.frx":15A841
      Top             =   1200
      Width           =   825
   End
   Begin VB.Image Image5 
      Height          =   5985
      Left            =   0
      Picture         =   "frmReservation.frx":16AE5B
      Top             =   1200
      Width           =   825
   End
End
Attribute VB_Name = "frmReservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim paraTrain1 As ADODB.Parameter
Dim paraTrain2 As ADODB.Parameter
Dim paraTrain As ADODB.Parameter
Dim rstTrainStn As ADODB.Recordset
Dim cmdTrain As ADODB.Command
Dim rstReserv As ADODB.Recordset
Dim cmdTrain1 As ADODB.Command
Dim rstTrainSeat As ADODB.Recordset
Dim rstTrainSeat1 As ADODB.Recordset
Dim rstRoute As ADODB.Recordset
'Dim rstSumDist As ADODB.Recordset
Dim rstFare As ADODB.Recordset
Dim stopageSource1 As Integer
Dim stopageDest1 As Integer
Dim stopageSource As Integer
Dim stopageDest As Integer
Dim stopageSource2 As Integer
Dim stopageDest2 As Integer
Dim count1 As Integer
Dim temp1 As Integer
Dim temp2 As Variant
Public fare As Currency
Dim sid As Integer
Public pnr As Variant
Private Function checkFields() As Boolean
count1 = 0
    For i = 0 To 4
        If txtName(i) <> "" Then
            If txtAge(i) = "" Or cmbGender(i) = "" Then
                checkFields = False
                Exit Function
            End If
            count1 = count1 + 1
        End If
    Next
    If count1 > 0 Then
        checkFields = True
    Else
        checkFields = False
    End If
End Function
Private Function checkStopage() As Boolean
    
    Set cmdTrain = New ADODB.Command
    cmdTrain.CommandType = adCmdTable
    cmdTrain.CommandText = "query4"
    cmdTrain.ActiveConnection = railCn
    Set paraTrain2 = cmdTrain.CreateParameter("seatNo", adVariant, adParamInput)
    cmdTrain.Parameters.Append paraTrain2
    paraTrain2.Value = frmSeatAvailablity.trainNo
    Set paraTrain = cmdTrain.CreateParameter("coachname", adVariant, adParamInput)
    cmdTrain.Parameters.Append paraTrain
    paraTrain.Value = temp2
    Set paraTrain1 = cmdTrain.CreateParameter("seatNo", adInteger, adParamInput)
    cmdTrain.Parameters.Append paraTrain1
    paraTrain1.Value = temp1
    Set rstTrainSeat1 = cmdTrain.Execute
    rstTrainSeat1.MoveFirst
    Set rstTrainStn = New ADODB.Recordset
    rstTrainStn.CursorLocation = adUseClient
    rstTrainStn.Open "select * from trainRoute where trainno='" & frmSeatAvailablity.trainNo & "' and (stnID =" & rstTrainSeat1("fromStn") & " or stnID =" & rstTrainSeat1("toStn") & ")order by stopageno", railCn
    stopageSource1 = rstTrainStn("stopageno")
    rstTrainStn.MoveNext
    stopageDest1 = rstTrainStn("stopageno")
    rstTrainStn.Close

    Set rstTrainStn = New ADODB.Recordset
    rstTrainStn.CursorLocation = adUseClient
    rstTrainStn.Open "select * from trainRoute where trainno='" & frmSeatAvailablity.trainNo & "' and (stnID =" & frmSeatAvailablity.sourceID & " or stnID =" & frmSeatAvailablity.destID & ") order by stopageno", railCn
    stopageSource = rstTrainStn("stopageno")
    rstTrainStn.MoveNext
    stopageDest = rstTrainStn("stopageno")
    rstTrainStn.Close

    If stopageSource >= stopageDest1 Then
        checkStopage = False
    Else
        checkStopage = True
    End If
End Function
Private Sub cmdBook_Click()
Dim count As Integer
Dim seat As Variant
Dim seatText As Variant
Dim temp As Integer
If (txtName(0).Text <> "" Or txtAge(0).Text <> "" Or cmbGender(0).Text <> "") And txtContact.Text <> "" Then
        If checkFields() Then
            If count1 <= frmSeatAvailablity.seatAvail Then
                
                Set rstFare = New ADODB.Recordset
                rstFare.CursorLocation = adUseClient
                rstFare.Open "select * from fare,train,coach where fare.typeid=train.traintypeid and fare.coachid=coach.coachtypeid and(uptrainno='" & frmSeatAvailablity.trainNo & "' or downtrainno='" & frmSeatAvailablity.trainNo & "') and coachid=" & frmSeatAvailablity.class & "", railCn
                If rstFare.RecordCount > 0 Then
                    fare = rstFare("fare")
                Else
                    MsgBox "Fare Information Not Exists For This Point Of Time.", vbCritical
                    Exit Sub
                End If
                Set rstRoute = New ADODB.Recordset
                rstRoute.CursorLocation = adUseClient
                rstRoute.Open "select * from routestn where routeid=" & rstFare("routeid") & " order by routeStnNo asc", railCn
                rstRoute.MoveFirst
                sid = rstRoute("stnid")
                rstRoute.Close
                If rstFare("upTrainNo") = trainNo Then
                    temp = 0
                    
                ElseIf rstFare("downTrainNo") = trainNo Then
                    temp = 1
                    
                End If
                If temp = 0 Then
                    dist1 = frmTrainSchedule.sumDist(sid, rstFare("routeid"), frmSeatAvailablity.sourceID)
                    dist2 = frmTrainSchedule.sumDist(sid, rstFare("routeid"), frmSeatAvailablity.destID)
                Else
                    dist1 = frmTrainSchedule.sumDist1(sid, rstFare("routeid"), frmSeatAvailablity.sourceID)
                    dist2 = frmTrainSchedule.sumDist1(sid, rstFare("routeid"), frmSeatAvailablity.destID)
                End If
                dist = dist2 - dist1
                fare = fare * count1 * dist
                Set cmdTrain1 = New ADODB.Command
                cmdTrain1.CommandType = adCmdText
                cmdTrain1.ActiveConnection = railCn
                Set rstReserv = New ADODB.Recordset
                rstReserv.CursorLocation = adUseClient
                rstReserv.Open "select max(pnrNo) from reservation", railCn
                If rstReserv.Fields(0) > 0 Then
                    pnr = rstReserv.Fields(0) + 1
                Else
                    pnr = 1111111111
                End If
                
                cmdTrain1.CommandText = "insert into reservation values('" & pnr & "','" & frmSeatAvailablity.trainNo & "',#" & Date & "#,#" & CDate(lblDateJourney.Caption) & "#," & frmSeatAvailablity.sourceID & " ," & frmSeatAvailablity.destID & ", '" & txtContact.Text & "'," & False & "," & fare & ")"
                cmdTrain1.Execute
                For i = 0 To 4
                
                    If txtName(i) <> "" Then
                            Set rstTrainSeat = New ADODB.Recordset
                            rstTrainSeat.CursorLocation = adUseClient
                            rstTrainSeat.Open "select * from trainseat,traincoach where trainseat.trainno=traincoach.trainno and coachtypeid=" & frmSeatAvailablity.class & " and journeydate=#" & lblDateJourney & "# and trainseat.trainno='" & frmSeatAvailablity.trainNo & "' and  trainseat.coachname=traincoach.coachname", railCn
                            If rstTrainSeat.RecordCount > 0 Then
                                n = 0
                                Do While (n <= rstTrainSeat.RecordCount - 1)
                                    k = 5
                                    For j = 0 To rstTrainSeat("totalSeat") - 1
                                        If Mid(rstTrainSeat("availableSeat"), k, 1) = "N" Then
                                            Set rstTrainStn = New ADODB.Recordset
                                            rstTrainStn.CursorLocation = adUseClient
                                            rstTrainStn.Open "select * from trainRoute where trainno='" & frmSeatAvailablity.trainNo & "' order by stopageno asc", railCn
                                            stopageSource1 = rstTrainStn("stopageno")
                                            rstTrainStn.MoveLast
                                            stopageDest1 = rstTrainStn("stopageno")
                                            rstTrainStn.Close
                                            
                                            Set rstTrainStn = New ADODB.Recordset
                                            rstTrainStn.CursorLocation = adUseClient
                                            rstTrainStn.Open "select * from trainRoute where trainno='" & frmSeatAvailablity.trainNo & "' and (stnID =" & frmSeatAvailablity.sourceID & " or stnID =" & frmSeatAvailablity.destID & ") order by stopageno asc", railCn
                                            If rstTrainStn.RecordCount > 1 Then
                                                stopageSource = rstTrainStn("stopageno")
                                                rstTrainStn.MoveNext
                                                stopageDest = rstTrainStn("stopageno")
                                                rstTrainStn.Close
                                                
                                                If stopageSource1 = stopageSource And stopageDest1 = stopageDest Then
                                                    seatText = Replace(rstTrainSeat("availableSeat"), Mid(rstTrainSeat("availableSeat"), k - 3, 4), Mid(rstTrainSeat("availableSeat"), k - 3, 3) & "B")
                                                    seat = Val(Mid(rstTrainSeat("availableSeat"), k - 3, 3))
                                                    Exit Do
                                                Else
                                                    seatText = Replace(rstTrainSeat("availableSeat"), Mid(rstTrainSeat("availableSeat"), k - 3, 4), Mid(rstTrainSeat("availableSeat"), k - 3, 3) & "P")
                                                    seat = Val(Mid(rstTrainSeat("availableSeat"), k - 3, 3))
                                                    Exit Do
                                                End If
                                            Else
                                                Label8.Visible = False
                                                MsgBox "Train Not Run Between Station That You Have Selected", vbCritical
                                            End If
                                        ElseIf Mid(rstTrainSeat("availableSeat"), k, 1) = "P" Then
                                                
                                            tempstr = Val(Mid(rstTrainSeat("availableSeat"), k - 3, 3))
                                            temp1 = tempstr
                                            temp2 = rstTrainSeat("trainseat.CoachName")
                                            Set cmdTrain = New ADODB.Command
                                            cmdTrain.CommandType = adCmdTable
                                            If checkStopage() Then
                                                cmdTrain.CommandText = "query4"
                                            Else
                                                cmdTrain.CommandText = "query5"
                                            End If
                                            cmdTrain.ActiveConnection = railCn
                                            Set paraTrain2 = cmdTrain.CreateParameter("seatNo", adVariant, adParamInput)
                                            cmdTrain.Parameters.Append paraTrain2
                                            paraTrain2.Value = frmSeatAvailablity.trainNo
                                            Set paraTrain = cmdTrain.CreateParameter("coachname", adVariant, adParamInput)
                                            cmdTrain.Parameters.Append paraTrain
                                            
                                            paraTrain.Value = rstTrainSeat("trainseat.CoachName")
                                            Set paraTrain1 = cmdTrain.CreateParameter("seatNo", adInteger, adParamInput)
                                            cmdTrain.Parameters.Append paraTrain1
                                            paraTrain1.Value = tempstr
                                            
                                            Set rstTrainSeat1 = cmdTrain.Execute
                                            rstTrainSeat1.MoveFirst
                                            
                                            For m = 0 To rstTrainSeat1.RecordCount - 1
                                                Set rstTrainStn = New ADODB.Recordset
                                                rstTrainStn.CursorLocation = adUseClient
                                                rstTrainStn.Open "select * from trainRoute where trainno='" & frmSeatAvailablity.trainNo & "' and (stnID =" & rstTrainSeat1("fromStn") & " or stnID =" & rstTrainSeat1("toStn") & ")order by stopageno", railCn
                                                stopageSource1 = rstTrainStn("stopageno")
                                                rstTrainStn.MoveNext
                                                stopageDest1 = rstTrainStn("stopageno")
                                                rstTrainStn.Close
                                                
                                                Set rstTrainStn = New ADODB.Recordset
                                                rstTrainStn.CursorLocation = adUseClient
                                                rstTrainStn.Open "select * from trainRoute where trainno='" & frmSeatAvailablity.trainNo & "' order by stopageno asc", railCn
                                                stopageSource2 = rstTrainStn("stopageno")
                                                rstTrainStn.MoveLast
                                                stopageDest2 = rstTrainStn("stopageno")
                                                rstTrainStn.Close

                                                Set rstTrainStn = New ADODB.Recordset
                                                rstTrainStn.CursorLocation = adUseClient
                                                rstTrainStn.Open "select * from trainRoute where trainno='" & frmSeatAvailablity.trainNo & "' and (stnID =" & frmSeatAvailablity.sourceID & " or stnID =" & frmSeatAvailablity.destID & ") order by stopageno", railCn
                                                If rstTrainStn.RecordCount > 1 Then
                                                    stopageSource = rstTrainStn("stopageno")
                                                    rstTrainStn.MoveNext
                                                    stopageDest = rstTrainStn("stopageno")
                                                    rstTrainStn.Close
                                                        If (stopageSource <= stopageSource1 And (stopageDest > stopageSource1 And stopageDest <= stopageDest1)) Or ((stopageSource >= stopageSource1 And stopageSource < stopageDest1) And stopageDest >= stopageDest1) Or (stopageSource < stopageSource1 And stopageDest > stopageDest1) Or (stopageSource > stopageSource1 And stopageDest < stopageDest1) Or (stopageSource = stopageSource1 And stopageDest = stopageDest1) Then
                                                            Exit For
                                                        End If
                                                        seatText = Replace(rstTrainSeat("availableSeat"), Mid(rstTrainSeat("availableSeat"), k - 3, 4), Mid(rstTrainSeat("availableSeat"), k - 3, 3) & "P")
                                                        seat = Val(Mid(rstTrainSeat("availableSeat"), k - 3, 3))
                                                        
                                                        Exit Do
                                                Else
                                                    Label8.Visible = False
                                                    MsgBox "Train Not Run Between Station That You Have Selected", vbCritical
                                                End If
                                                rstTrainSeat1.MoveNext
                                            Next

                                        End If
                                        k = k + 5
                                    Next
                                    n = n + 1
                                Loop
                            End If
                            

                        cmdTrain1.CommandText = "update trainSeat set availableSeat='" & seatText & "' where trainno='" & frmSeatAvailablity.trainNo & "' and journeydate=#" & CDate(lblDateJourney.Caption) & "# and coachname='" & rstTrainSeat("trainseat.coachname") & "'"
                        cmdTrain1.Execute
                        cmdTrain1.CommandText = "insert into passenger values('" & pnr & "','" & rstTrainSeat("trainseat.coachname") & "'," & seat & ",'" & txtName(i).Text & "'," & txtAge(i).Text & " ,'" & cmbGender(i).Text & "')"
                        cmdTrain1.Execute
                    End If
                    k = k + 5
                Next
                'MsgBox "Ticket is Successfully Reserved.", vbInformation
                frmTicketInfo.Show
                Unload Me
            Else
                MsgBox "Sorry,Seats is Not Available for " & count & "Passengers", vbCritical
            End If
        Else
            MsgBox "You Have To Fill Atleast One Passenger Information.", vbCritical
        End If
    Else
        MsgBox "Please Fill All Fields.", vbCritical
    End If
End Sub

Private Sub Form_load()
For i = 0 To 4
    cmbGender(i).AddItem "M"
    cmbGender(i).AddItem "F"
Next
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


Private Sub txtAge_KeyPress(Index As Integer, KeyAscii As Integer)
    Call validation(1, KeyAscii, txtAge(Index))
End Sub

Private Sub txtContact_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(txtContact.Text) < 10 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            txtContact.Text = txtContact.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        txtContact.Text = txtContact.Text & Chr(KeyAscii)
    End If
End If
End Sub


Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
    Call validation(2, KeyAscii, txtName(Index))
End Sub
