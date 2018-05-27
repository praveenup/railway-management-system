VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Indian Railways"
   ClientHeight    =   8970
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14550
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      Picture         =   "MDIForm1.frx":39E8F
      ScaleHeight     =   1905
      ScaleWidth      =   14520
      TabIndex        =   1
      Top             =   0
      Width           =   14550
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "Bernard MT Condensed"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   19800
         Picture         =   "MDIForm1.frx":E7DA1
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   330
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
         Left            =   5160
         Picture         =   "MDIForm1.frx":E9463
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   3015
      End
      Begin VB.CommandButton Command4 
         Height          =   1695
         Left            =   11160
         Picture         =   "MDIForm1.frx":FE8D1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   3015
      End
      Begin VB.CommandButton Command3 
         Height          =   1695
         Left            =   14160
         Picture         =   "MDIForm1.frx":114BA3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Height          =   1695
         Left            =   8160
         Picture         =   "MDIForm1.frx":12596D
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   3015
      End
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
         Left            =   2160
         Picture         =   "MDIForm1.frx":136737
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   3015
      End
      Begin VB.Line Line3 
         BorderWidth     =   15
         X1              =   17520
         X2              =   20400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line2 
         BorderWidth     =   16
         X1              =   17520
         X2              =   20400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderWidth     =   15
         X1              =   17415
         X2              =   17415
         Y1              =   0
         Y2              =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "l3"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   18120
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "l2"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   18240
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.Image Image5 
         Height          =   720
         Left            =   17520
         Picture         =   "MDIForm1.frx":147501
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H000000C0&
      Height          =   7035
      Left            =   0
      Picture         =   "MDIForm1.frx":147C42
      ScaleHeight     =   6975
      ScaleWidth      =   2310
      TabIndex        =   0
      Top             =   1935
      Width           =   2370
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   0
         Picture         =   "MDIForm1.frx":1C9970
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CommandButton Command7 
         Height          =   375
         Left            =   0
         Picture         =   "MDIForm1.frx":1CCB0E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   0
         Picture         =   "MDIForm1.frx":1CFCAC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Section"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   0
         Width           =   2295
      End
      Begin VB.Image Image3 
         Height          =   420
         Left            =   0
         Picture         =   "MDIForm1.frx":1D2E4A
         Top             =   0
         Width           =   11535
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnuTrainInfo 
         Caption         =   "Train Information And Route"
      End
      Begin VB.Menu mnuTrainCoaches 
         Caption         =   "Train Coaches"
      End
      Begin VB.Menu mnuRoute 
         Caption         =   "Route"
      End
      Begin VB.Menu mnuCoachType 
         Caption         =   "Coach Type"
      End
      Begin VB.Menu mnuFare 
         Caption         =   "Fare"
      End
      Begin VB.Menu mnuCoaches 
         Caption         =   "Coaches"
      End
      Begin VB.Menu mnuStation 
         Caption         =   "Station"
      End
      Begin VB.Menu mnuTrainType 
         Caption         =   "Train Type"
      End
      Begin VB.Menu mnuBerth 
         Caption         =   "Berth"
      End
      Begin VB.Menu mnuQuota 
         Caption         =   "Quota"
      End
      Begin VB.Menu mnuRegion 
         Caption         =   "Region"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuReservation 
         Caption         =   "Reservation"
      End
      Begin VB.Menu mnuCancellation 
         Caption         =   "Cancellation"
      End
   End
   Begin VB.Menu mnuEnquiry 
      Caption         =   "Train &Enquiry"
      Begin VB.Menu mnuPnr 
         Caption         =   "PNR Status"
      End
      Begin VB.Menu mnuTrainBet 
         Caption         =   "Train Between Stn"
      End
      Begin VB.Menu mnuSchedule 
         Caption         =   "Train Schedule"
      End
      Begin VB.Menu mnuSeatAvail 
         Caption         =   "Seat Availablity"
      End
      Begin VB.Menu mnuFareEnquiry 
         Caption         =   "Fare Enquiry"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstUser As ADODB.Recordset
Dim rstSeat As ADODB.Recordset
Dim rstTrain As ADODB.Recordset
Dim rstCoach As ADODB.Recordset
Dim cmdSeat As ADODB.Command

Private Sub Command1_Click()
frmPnr.Show
End Sub

Private Sub Command2_Click()
frmTrainSchedule.Show
End Sub



Private Sub Command3_Click()
frmFareEnquiry.Show
End Sub

Private Sub Command4_Click()
frmSeatAvailablity.Show
End Sub

Private Sub Command5_Click()
frmTrainBetStn.Show
End Sub

Private Sub Command6_Click()
frmSeatAvailablity.Show
End Sub

Private Sub Command7_Click()
frmCancellation.Show
End Sub

Private Sub MDIForm_Load()
    Set rstUser = New ADODB.Recordset
    rstUser.CursorLocation = adUseClient
    rstUser.Open "select * from useraccount where accountid=" & userAccountID & "", railCn
    If rstUser.RecordCount > 0 Then
        Label2.Caption = "Hi " & rstUser("username")
        Label3.Caption = rstUser("usertype")
    End If
    rstUser.Close
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim strSeat As Variant
Set rstTrain = New ADODB.Recordset
rstTrain.CursorLocation = adUseClient
rstTrain.Open "select uptrainno,downtrainno from train", railCn
If rstTrain.RecordCount > 0 Then
    rstTrain.MoveFirst
    For j = 0 To rstTrain.RecordCount - 1
        Set rstCoach = New ADODB.Recordset
        rstCoach.CursorLocation = adUseClient
        rstCoach.Open "select coachname,cabin,cabinseats from traincoach,coach where trainno='" & rstTrain(0) & "'  and coach.coachtypeid=traincoach.coachtypeid and passengercoach=" & True & "", railCn
        If rstCoach.RecordCount > 0 Then
            rstCoach.MoveFirst
            Dim temp As Integer
            For k = 0 To rstCoach.RecordCount - 1
                temp = rstCoach(1) * rstCoach(2)
                strSeat = ""
                For i = 1 To temp
                    If strSeat = "" Then
                        strSeat = "-00" & i & "N-"
                    ElseIf i < 10 Then
                        If i = temp Then
                            strSeat = strSeat & "00" & i & "N"
                        Else
                            strSeat = strSeat & "00" & i & "N-"
                        End If
                    ElseIf i >= 10 And i < 100 Then
                        If i = temp Then
                            strSeat = strSeat & "0" & i & "N"
                        Else
                            strSeat = strSeat & "0" & i & "N-"
                        End If
                    Else
                        If i = temp Then
                            strSeat = strSeat & i & "N"
                        Else
                            strSeat = strSeat & i & "N-"
                        End If
                    End If

                Next
                For i = 0 To 9
                    Set rstSeat = New ADODB.Recordset
                    rstSeat.CursorLocation = adUseClient
                    rstSeat.Open "select journeydate from trainseat where trainno='" & rstTrain(0) & "' and journeydate = #" & CDate(DateAdd("d", i, Date)) & "# and coachname='" & rstCoach(0) & "'  order by journeydate asc", railCn

                    If rstSeat.RecordCount = 0 Then
'                        tempDate = CDate(DateAdd("d", i, Date))
'                        'If rstSeat(0) <> CDate(Date) Then
'                            If DateAdd("d", i, Date) <> rstSeat(0) Then
'                                Set cmdSeat = New ADODB.Command
'                                cmdSeat.CommandType = adCmdText
'                                cmdSeat.ActiveConnection = railCn
'                                cmdSeat.CommandText = "insert into trainSeat values('" & rstTrain(0) & "',#" & tempDate & "#,'" & rstCoach(0) & "','" & strSeat & "'," & temp & ")"
'                                cmdSeat.Execute
'                            End If
'                        Else
'                            Set cmdSeat = New ADODB.Command
'                            cmdSeat.CommandType = adCmdText
'                            cmdSeat.ActiveConnection = railCn
'                            cmdSeat.CommandText = "insert into trainSeat values('" & rstTrain(0) & "','" & Date & "','" & rstCoach(0) & "','" & strSeat & "'," & temp & ")"
'                            cmdSeat.Execute
                        'End If
                    'Else
                        Set cmdSeat = New ADODB.Command
                        cmdSeat.CommandType = adCmdText
                        cmdSeat.ActiveConnection = railCn
                        cmdSeat.CommandText = "insert into trainSeat values('" & rstTrain(0) & "','" & CDate(DateAdd("d", i, Date)) & "','" & rstCoach(0) & "','" & strSeat & "'," & temp & ")"
                        cmdSeat.Execute
                    End If
                    rstSeat.Close
                Next
            rstCoach.MoveNext
            Next
        End If

        Set rstCoach = New ADODB.Recordset
        rstCoach.CursorLocation = adUseClient
        rstCoach.Open "select coachname,cabin,cabinseats from traincoach,coach where trainno='" & rstTrain(1) & "'  and coach.coachtypeid=traincoach.coachtypeid and passengercoach=" & True & "", railCn
        If rstCoach.RecordCount > 0 Then
            rstCoach.MoveFirst
            'Dim temp As Integer
            For k = 0 To rstCoach.RecordCount - 1
                temp = rstCoach(1) * rstCoach(2)
                strSeat = ""
                For i = 1 To temp
                    If strSeat = "" Then
                        strSeat = "-00" & i & "N-"
                    ElseIf i < 10 Then
                        If i = temp Then
                            strSeat = strSeat & "00" & i & "N"
                        Else
                            strSeat = strSeat & "00" & i & "N-"
                        End If
                    ElseIf i >= 10 And i < 100 Then
                        If i = temp Then
                            strSeat = strSeat & "0" & i & "N"
                        Else
                            strSeat = strSeat & "0" & i & "N-"
                        End If
                    Else
                        If i = temp Then
                            strSeat = strSeat & i & "N"
                        Else
                            strSeat = strSeat & i & "N-"
                        End If
                    End If

                Next
                For i = 0 To 9
                    Set rstSeat = New ADODB.Recordset
                    rstSeat.CursorLocation = adUseClient
                    rstSeat.Open "select journeydate from trainseat where trainno='" & rstTrain(1) & "' and journeydate = #" & CDate(DateAdd("d", i, Date)) & "# and coachname='" & rstCoach(0) & "'  order by journeydate asc", railCn
'                    If checkDate() Then
'
'                    Else
'
'                    End If
                    If rstSeat.RecordCount = 0 Then
'                        tempDate = CDate(DateAdd("d", i, Date))
'                        'If rstSeat(0) <> CDate(Date) Then
'                            If DateAdd("d", i, Date) <> rstSeat(0) Then
'                                Set cmdSeat = New ADODB.Command
'                                cmdSeat.CommandType = adCmdText
'                                cmdSeat.ActiveConnection = railCn
'                                cmdSeat.CommandText = "insert into trainSeat values('" & rstTrain(1) & "','" & tempDate & "','" & rstCoach(0) & "','" & strSeat & "'," & temp & ")"
'                                cmdSeat.Execute
'                            End If
'                        'End If
'                    Else
                        Set cmdSeat = New ADODB.Command
                        cmdSeat.CommandType = adCmdText
                        cmdSeat.ActiveConnection = railCn
                        cmdSeat.CommandText = "insert into trainSeat values('" & rstTrain(1) & "','" & CDate(DateAdd("d", i, Date)) & "','" & rstCoach(0) & "','" & strSeat & "'," & temp & ")"
                        cmdSeat.Execute
                    End If
                    rstSeat.Close
                Next
            rstCoach.MoveNext
            Next
        End If
        rstTrain.MoveNext
    Next
End If
End Sub

Private Sub mnuBerth_Click()
    frmBerth.Show
End Sub

Private Sub mnuCancellation_Click()
frmCancellation.Show
End Sub

Private Sub mnuCoaches_Click()
    frmCoaches.Show
End Sub

Private Sub mnuCoachType_Click()
    frmCoachType.Show
End Sub

Private Sub mnuFare_Click()
    frmFare.Show
End Sub

Private Sub mnuPnr_Click()
frmPnr.Show
End Sub

Private Sub mnuQuota_Click()
    frmQuota.Show
End Sub

Private Sub mnuRegion_Click()
    frmRegion.Show
End Sub

Private Sub mnuReservation_Click()
frmSeatAvailablity.Show
End Sub

Private Sub mnuRoute_Click()
    frmRoute.Show
End Sub

Private Sub mnuSchedule_Click()
frmTrainSchedule.Show
End Sub

Private Sub mnuSeatAvail_Click()
frmSeatAvailablity.Show
End Sub

Private Sub mnuStation_Click()
    frmStation.Show
End Sub

Private Sub mnuTrainBet_Click()
frmTrainBetStn.Show
End Sub

Private Sub mnuTrainCoaches_Click()
frmTrainCoach.Show
End Sub

Private Sub mnuTrainInfo_Click()
    frmTrain.Show
End Sub

Private Sub mnuTrainType_Click()
    frmTrainType.Show
End Sub



