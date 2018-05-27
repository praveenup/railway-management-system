VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSeatAvailablity 
   BackColor       =   &H8000000E&
   Caption         =   "Seat Availability"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15855
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   15855
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   480
   End
   Begin VB.CommandButton cmdBook 
      Height          =   375
      Left            =   6960
      Picture         =   "frmSeatAvailablity.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7200
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   11760
      TabIndex        =   13
      Top             =   7920
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   12120
      TabIndex        =   12
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Fill All Fields"
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
      Height          =   6015
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   14775
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   3600
         Picture         =   "frmSeatAvailablity.frx":4356
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtTrainNo 
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cmbClass 
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
         Left            =   11160
         TabIndex        =   4
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtSource 
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
         Left            =   3240
         TabIndex        =   3
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtDest 
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
         Left            =   9720
         TabIndex        =   2
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton cmdCheck 
         Height          =   375
         Left            =   6000
         Picture         =   "frmSeatAvailablity.frx":51F0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3480
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dateJourney 
         Height          =   375
         Left            =   6600
         TabIndex        =   5
         Top             =   1320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108789761
         CurrentDate     =   42527
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Seat Availablity Details"
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
         Left            =   -360
         TabIndex        =   18
         Top             =   4200
         Width           =   15975
      End
      Begin VB.Image Image7 
         Height          =   405
         Left            =   0
         Picture         =   "frmSeatAvailablity.frx":9546
         Top             =   4200
         Width           =   16005
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Fill All Fields Of Your Train Journey Details"
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
         Left            =   -360
         TabIndex        =   16
         Top             =   0
         Width           =   15975
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   0
         Picture         =   "frmSeatAvailablity.frx":1E774
         Top             =   0
         Width           =   16005
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Train Details"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   975
         Left            =   3960
         TabIndex        =   14
         Top             =   4680
         Width           =   6735
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
         TabIndex        =   11
         Top             =   1440
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
         Left            =   4680
         TabIndex        =   10
         Top             =   1440
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
         Left            =   10200
         TabIndex        =   9
         Top             =   1320
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
         Left            =   1440
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
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
         Left            =   8280
         TabIndex        =   7
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Image Image8 
         Height          =   5925
         Left            =   0
         Picture         =   "frmSeatAvailablity.frx":339A2
         Top             =   0
         Width           =   14790
      End
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   360
      Picture         =   "frmSeatAvailablity.frx":151114
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Seat Availability"
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
      Left            =   1200
      TabIndex        =   17
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmSeatAvailablity.frx":151A5F
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11520
      Picture         =   "frmSeatAvailablity.frx":151F39
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   15720
      Picture         =   "frmSeatAvailablity.frx":152413
      Top             =   840
      Width           =   825
   End
   Begin VB.Image Image5 
      Height          =   5985
      Left            =   120
      Picture         =   "frmSeatAvailablity.frx":162A2D
      Top             =   840
      Width           =   825
   End
End
Attribute VB_Name = "frmSeatAvailablity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCoach As ADODB.Recordset
Dim rstDay As ADODB.Recordset
Dim paraTrain1 As ADODB.Parameter
Dim paraTrain2 As ADODB.Parameter
Dim paraTrain As ADODB.Parameter
Dim rstSourceStn As ADODB.Recordset
Dim rstDestStn As ADODB.Recordset
Dim rstTrain As ADODB.Recordset
Dim rstTrain1 As ADODB.Recordset
Dim rstTrainSeat As ADODB.Recordset
Dim rstTrainSeat1 As ADODB.Recordset
Dim rstTrainStn As ADODB.Recordset
Dim cmdTrain As ADODB.Command
Public sourceID As Integer
Public destID As Integer
Dim stopageSource As Integer
Dim stopageDest As Integer
Dim stopageSource1 As Integer
Dim stopageDest1 As Integer
Dim trainName As Variant
Public seatAvail As Integer
Public trainNo As Variant
Public class As Integer
Dim flag As Integer
Dim temp1 As Integer
Dim temp2 As Variant
Private Function checkTrainNo() As Boolean
    Set rstTrain = New ADODB.Recordset
    rstTrain.CursorLocation = adUseClient
    rstTrain.Open "select * from train ", railCn
    If rstTrain.RecordCount > 0 Then
        rstTrain.MoveFirst
        For i = 0 To rstTrain.RecordCount - 1
            If rstTrain("uptrainno") = txtTrainNo.Text Or rstTrain("downtrainno") = txtTrainNo.Text Then
                checkTrainNo = True
                rstTrain.Close
                Exit Function
            End If
            rstTrain.MoveNext
        Next
    End If
    checkTrainNo = False
    rstTrain.Close
End Function
'Private Sub cmdNext_Click()
''Frame1.Height = 7575
'If destID <> sourceID And txtTrainNo.Text <> "" And dateJourney >= Date And dateJourney <= DateAdd("d", 9, Date) And txtDest.Text <> "" And txtSource.Text <> "" And cmbClass.ListIndex <> -1 Then
'
'End If
'End Sub

Private Sub cmdBook_Click()
If seatAvail > 0 Then
    If flag = 1 Then
        If destID <> sourceID And txtTrainNo.Text <> "" And dateJourney >= Date And dateJourney <= DateAdd("d", 9, Date) And txtDest.Text <> "" And txtSource.Text <> "" And cmbClass.ListIndex <> -1 Then
            frmReservation.lblTrainNo.Caption = trainName & "(" & txtTrainNo.Text & ")"
            frmReservation.lblDateJourney.Caption = dateJourney.Value
            frmReservation.lblClass.Caption = cmbClass.Text
            frmReservation.lblFromStn.Caption = txtSource.Text
            frmReservation.lblToStn.Caption = txtDest.Text
            trainNo = txtTrainNo.Text
            frmReservation.Show
            Unload Me
        Else
            MsgBox "First Check Seat Availablity.", vbCritical
        End If
    Else
        MsgBox "First Check Seat Availablity.", vbCritical
    End If
Else
    MsgBox "Seat Availablity Zero.", vbCritical
End If
End Sub
Private Function checkStopage() As Boolean
    
    Set cmdTrain = New ADODB.Command
    cmdTrain.CommandType = adCmdTable
    cmdTrain.CommandText = "query4"
    cmdTrain.ActiveConnection = railCn
    Set paraTrain2 = cmdTrain.CreateParameter("seatNo", adVariant, adParamInput)
    cmdTrain.Parameters.Append paraTrain2
    paraTrain2.Value = txtTrainNo.Text
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
    rstTrainStn.Open "select * from trainRoute where trainno='" & txtTrainNo.Text & "' and (stnID =" & rstTrainSeat1("fromStn") & " or stnID =" & rstTrainSeat1("toStn") & ")order by stopageno", railCn
    stopageSource1 = rstTrainStn("stopageno")
    rstTrainStn.MoveNext
    stopageDest1 = rstTrainStn("stopageno")
    rstTrainStn.Close

    Set rstTrainStn = New ADODB.Recordset
    rstTrainStn.CursorLocation = adUseClient
    rstTrainStn.Open "select * from trainRoute where trainno='" & txtTrainNo.Text & "' and (stnID =" & frmSeatAvailablity.sourceID & " or stnID =" & frmSeatAvailablity.destID & ") order by stopageno", railCn
    If rstTrainStn.RecordCount > 1 Then
        stopageSource = rstTrainStn("stopageno")
        rstTrainStn.MoveNext
        stopageDest = rstTrainStn("stopageno")
        rstTrainStn.Close
    
        If stopageSource >= stopageDest1 Then
            checkStopage = False
        Else
            checkStopage = True
        End If
    End If
End Function

Private Function checkDays() As Boolean
    Set rstDay = New ADODB.Recordset
    rstDay.Open "select * from days where trainno='" & txtTrainNo.Text & "'", railCn, 3, 3
    If rstDay.RecordCount > 0 Then
        For i = 0 To rstDay.RecordCount - 1
        'MsgBox dateJourney.DayOfWeek
            If True = rstDay(dateJourney.DayOfWeek) Then
                checkDays = True
                Exit Function
            End If
        Next
    End If
    checkDays = False
End Function
Private Sub cmdCheck_Click()
stopageSource = 0
stopageDest = 0
If destID <> sourceID And txtTrainNo.Text <> "" And dateJourney >= Date And dateJourney <= DateAdd("d", 9, Date) And txtDest.Text <> "" And txtSource.Text <> "" And cmbClass.ListIndex <> -1 Then
    If dateJourney <> Date Then
        If checkDays() Then
            Dim k As Integer
            Dim count As Integer
            Set rstTrainSeat = New ADODB.Recordset
            rstTrainSeat.CursorLocation = adUseClient
            rstTrainSeat.Open "select * from trainseat,traincoach where trainseat.trainno=traincoach.trainno and coachtypeid=" & cmbClass.ItemData(cmbClass.ListIndex) & " and journeydate=#" & CDate(dateJourney.Value) & "# and trainseat.trainno='" & txtTrainNo.Text & "' and  trainseat.coachname=traincoach.coachname", railCn
            If rstTrainSeat.RecordCount > 0 Then
                For i = 0 To rstTrainSeat.RecordCount - 1
                    k = 5
                    For j = 0 To rstTrainSeat("totalSeat") - 1
                        If Mid(rstTrainSeat("availableSeat"), k, 1) = "N" Then
                            Set rstTrainStn = New ADODB.Recordset
                            rstTrainStn.CursorLocation = adUseClient
                            rstTrainStn.Open "select * from trainRoute where trainno='" & txtTrainNo & "' and stnID =" & sourceID & "", railCn
                            If rstTrainStn.RecordCount > 0 Then
                                stopageSource = rstTrainStn("stopageno")
                            End If
                            
                            Set rstTrainStn = New ADODB.Recordset
                            rstTrainStn.CursorLocation = adUseClient
                            rstTrainStn.Open "select * from trainRoute where trainno='" & txtTrainNo & "' and  stnID =" & destID & "", railCn
                            If rstTrainStn.RecordCount > 0 Then
                                stopageDest = rstTrainStn("stopageno")
                            End If
                            rstTrainStn.Close
                            If stopageSource > stopageDest Or stopageSource = 0 Or stopageDest = 0 Then
                                Label8.Visible = False
                                Label7.Visible = False
                                Image7.Visible = False
                                cmdBook.Visible = False
                                MsgBox "Train Not Run Between Station That You Have Selected", vbCritical
                                Exit Sub
                            End If
                            count = count + 1
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
                            paraTrain2.Value = txtTrainNo.Text
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
                                rstTrainStn.Open "select * from trainRoute where trainno='" & txtTrainNo & "' and (stnID =" & rstTrainSeat1("fromStn") & " or stnID =" & rstTrainSeat1("toStn") & ") order by stopageno", railCn
                                stopageSource1 = rstTrainStn("stopageno")
                                rstTrainStn.MoveNext
                                stopageDest1 = rstTrainStn("stopageno")
                                rstTrainStn.Close
                                
                                Set rstTrainStn = New ADODB.Recordset
                                rstTrainStn.CursorLocation = adUseClient
                                rstTrainStn.Open "select * from trainRoute where trainno='" & txtTrainNo & "' and stnID =" & sourceID & "", railCn
                                If rstTrainStn.RecordCount > 0 Then
                                    stopageSource = rstTrainStn("stopageno")
                                End If
                                
                                Set rstTrainStn = New ADODB.Recordset
                                rstTrainStn.CursorLocation = adUseClient
                                rstTrainStn.Open "select * from trainRoute where trainno='" & txtTrainNo & "' and  stnID =" & destID & "", railCn
                                If rstTrainStn.RecordCount > 0 Then
                                    stopageDest = rstTrainStn("stopageno")
                                End If
                                rstTrainStn.Close
                                    If stopageSource > stopageDest Or stopageSource = 0 Or stopageDest = 0 Then
                                        Label8.Visible = False
                                        Label7.Visible = False
                                        Image7.Visible = False
                                        cmdBook.Visible = False
                                        MsgBox "Train Not Run Between Station That You Have Selected", vbCritical
                                        Exit Sub
                                    'ElseIf (stopageSource1 >= stopageSource And stopageDest1 > stopageDest) Or (stopageSource1 < stopageSource And stopageDest1 <= stopageDest) Then
                                     ElseIf Not ((stopageSource <= stopageSource1 And (stopageDest > stopageSource1 And stopageDest <= stopageDest1)) Or ((stopageSource >= stopageSource1 And stopageSource < stopageDest1) And stopageDest >= stopageDest1) Or (stopageSource < stopageSource1 And stopageDest > stopageDest1) Or (stopageSource > stopageSource1 And stopageDest < stopageDest1) Or (stopageSource = stopageSource1 And stopageDest = stopageDest1)) Then
                                        count = count + 1
                                        Exit For
                                    End If
                                    Exit For
    '
                                rstTrainSeat1.MoveNext
                            Next
                        End If
                        k = k + 5
                    Next
                    rstTrainSeat.MoveNext
                Next
                Set rstTrain1 = New ADODB.Recordset
                rstTrain1.CursorLocation = adUseClient
                rstTrain1.Open "select * from train where Uptrainno='" & txtTrainNo.Text & "' or downtrainno='" & txtTrainNo.Text & "'", railCn
                Label8.Visible = True
                Label7.Visible = True
                Image7.Visible = True
                trainName = rstTrain1("trainname")
                Label8.Caption = "Train Name : " & rstTrain1("trainname") & "(" & txtTrainNo.Text & ")" & Chr(vbKeyReturn) & "Class Type : " & cmbClass.Text & Chr(vbKeyReturn) & "Available Seats : " & count
                seatAvail = count
                class = cmbClass.ItemData(cmbClass.ListIndex)
                Label8.Height = 975
                rstTrain1.Close
                flag = 1
                cmdBook.Visible = True
            Else
                MsgBox "Train Does Not Have Class That You Have Selected.", vbCritical
                Label7.Visible = False
                Image7.Visible = False
                Label8.Visible = False
                cmdBook.Visible = False
            End If
        Else
            MsgBox "Train Not Run in Date " & dateJourney.Value, vbCritical
            Label7.Visible = False
            Image7.Visible = False
            Label8.Visible = False
            cmdBook.Visible = False
        End If
    Else
        MsgBox "Ticket Cannot Be Booked in Date Of Train Running", vbCritical
        Label7.Visible = False
        Image7.Visible = False
        Label8.Visible = False
        cmdBook.Visible = False
    End If
Else
    MsgBox "Incorrect Input Information", vbCritical
    Label7.Visible = False
    Image7.Visible = False
    Label8.Visible = False
    cmdBook.Visible = False
End If
End Sub



Private Sub cmdSearch_Click()
frmSearchDialog.Show 1
End Sub

Private Sub Form_load()
    Set rstCoach = New ADODB.Recordset
    rstCoach.CursorLocation = adUseClient
    rstCoach.Open "select * from coach", railCn
    If rstCoach.RecordCount > 0 Then
        i = 0
        rstCoach.MoveFirst
        Do While Not rstCoach.EOF
            cmbClass.AddItem rstCoach(1)
            cmbClass.ItemData(i) = rstCoach(0)
            rstCoach.MoveNext
            i = i + 1
        Loop
    End If
    rstCoach.Close
    

    
    List1.Visible = False
    List1.Top = Frame1.Top + txtSource.Top + txtSource.Height
    List1.Left = txtSource.Left + Frame1.Left
    List2.Visible = False
    List2.Top = Frame1.Top + txtDest.Top + txtDest.Height
    List2.Left = txtDest.Left + Frame1.Left
    Label7.Visible = False
    Image7.Visible = False
    Label8.Visible = False
    cmdBook.Visible = False
    dateJourney.Value = Date
'    Frame1.Height = 2200
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

Private Sub txtDest_Change()
    If txtDest.Text <> "" Then
        List2.Visible = True
        List2.Clear
        Set rstDestStn = New ADODB.Recordset
        rstDestStn.CursorLocation = adUseClient
        rstDestStn.Open "select * from station where stnname like '" & "%" & txtDest.Text & "%'", railCn
        If rstDestStn.RecordCount > 0 Then
            i = 0
            rstDestStn.MoveFirst
            Do While Not rstDestStn.EOF
                List2.AddItem rstDestStn(2) & "(" & rstDestStn(1) & ")"
                List2.ItemData(i) = rstDestStn(0)
                rstDestStn.MoveNext
                i = i + 1
            Loop
        Else
            List2.Clear
            List2.Visible = False
        End If
        rstDestStn.Close
    Else
        List2.Visible = False
    End If
    
    'destID = 0
End Sub
Private Sub txtsource_Change()
    If txtSource.Text <> "" Then
        List1.Visible = True
        List1.Clear
        Set rstSourceStn = New ADODB.Recordset
        rstSourceStn.CursorLocation = adUseClient
        rstSourceStn.Open "select * from station where stnname like '" & "%" & txtSource.Text & "%'", railCn
        If rstSourceStn.RecordCount > 0 Then
            i = 0
            rstSourceStn.MoveFirst
            Do While Not rstSourceStn.EOF
                List1.AddItem rstSourceStn(2) & "(" & rstSourceStn(1) & ")"
                List1.ItemData(i) = rstSourceStn(0)
                rstSourceStn.MoveNext
                i = i + 1
            Loop
        Else
            List1.Clear
            List1.Visible = False
        End If
        rstSourceStn.Close
    Else
        List1.Visible = False
    End If
    'sourceID = 0
End Sub
Private Sub List1_Click()
sourceID = List1.ItemData(List1.ListIndex)
txtSource = List1.Text
List1.Visible = False
End Sub
Private Sub txtsource_LostFocus()
List1.Visible = False
End Sub
Private Sub list2_Click()
destID = List2.ItemData(List2.ListIndex)
txtDest = List2.Text
List2.Visible = False
End Sub
Private Sub txtdest_LostFocus()
List2.Visible = False
End Sub

Private Sub txtTrainNo_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(txtTrainNo.Text) < 5 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            txtTrainNo.Text = txtTrainNo.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        txtTrainNo.Text = txtTrainNo.Text & Chr(KeyAscii)
    End If
End If
End Sub
Private Sub txtTrainNo_LostFocus()
    If txtTrainNo.Text <> "" Then
        If checkTrainNo() = False Then
            MsgBox "TrainNo is Incorrect", vbCritical
            txtTrainNo.SetFocus
            txtTrainNo.Text = ""
        End If
    End If
End Sub
