VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFareEnquiryUser 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fare Enquiry"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   15570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   360
      TabIndex        =   15
      Top             =   0
      Width           =   2895
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
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Fill Below Fields To Find Fare Details"
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
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   13695
      Begin VB.CommandButton cmdGetFare 
         Height          =   615
         Left            =   6120
         Picture         =   "frmFareEnquiryUser.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2760
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   7920
         Picture         =   "frmFareEnquiryUser.frx":32A6
         ScaleHeight     =   2025
         ScaleWidth      =   5505
         TabIndex        =   6
         Top             =   2040
         Width           =   5535
         Begin VB.Label lblFare 
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
            Height          =   3735
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   5295
         End
      End
      Begin VB.ComboBox cmbCoach 
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
         ItemData        =   "frmFareEnquiryUser.frx":120A18
         Left            =   2880
         List            =   "frmFareEnquiryUser.frx":120A1A
         TabIndex        =   4
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtTrainNo 
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
         TabIndex        =   3
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtSource 
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
         TabIndex        =   2
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox txtDest 
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
         TabIndex        =   1
         Top             =   3840
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dateJourney 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   4800
         Width           =   2895
         _ExtentX        =   5106
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
         Format          =   111083521
         CurrentDate     =   42527
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Coach Type:"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "From Stn:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1080
         TabIndex        =   11
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Train No.:"
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
         Left            =   1080
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "To Stn:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1080
         TabIndex        =   9
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Journey Date:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1080
         TabIndex        =   8
         Top             =   4800
         Width           =   2295
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Fare enquiry"
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
      TabIndex        =   13
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   360
      Picture         =   "frmFareEnquiryUser.frx":120A1C
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmFareEnquiryUser.frx":121367
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11520
      Picture         =   "frmFareEnquiryUser.frx":121841
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image5 
      Height          =   5985
      Left            =   120
      Picture         =   "frmFareEnquiryUser.frx":121D1B
      Top             =   1320
      Width           =   825
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   14640
      Picture         =   "frmFareEnquiryUser.frx":132335
      Top             =   1320
      Width           =   825
   End
End
Attribute VB_Name = "frmFareEnquiryUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCoach As ADODB.Recordset
Dim rstTrain As ADODB.Recordset
Dim rstDestStn As ADODB.Recordset
Dim rstSourceStn As ADODB.Recordset
Dim rstFare As ADODB.Recordset
Dim rstRoute As ADODB.Recordset
Dim sourceID As Integer
Dim destID As Integer
Dim sid As Integer
Dim trainName As Variant
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
Private Sub cmdGetFare_Click()
If txtTrainNo.Text <> "" And dateJourney.Value <> 0 And sourceID <> 0 And destID <> 0 And cmbCoach.ListIndex <> -1 Then
    If checkDays() Then
        Set rstFare = New ADODB.Recordset
        rstFare.CursorLocation = adUseClient
        rstFare.Open "select * from fare,train,coach where fare.typeid=train.traintypeid and fare.coachid=coach.coachtypeid and(uptrainno='" & txtTrainNo & "' or downtrainno='" & txtTrainNo & "') and coachid=" & cmbCoach.ItemData(cmbCoach.ListIndex) & "", railCn
        If rstFare.RecordCount > 0 Then
            fare = rstFare("fare")
            trainName = rstFare("trainname")
        Else
            MsgBox "Fare Information Not Exists For This Point Of Time.", vbCritical
            Frame1.Width = 7815
            Image4.Left = 8760
            Exit Sub
        End If
        
        Set rstRoute = New ADODB.Recordset
        rstRoute.CursorLocation = adUseClient
        rstRoute.Open "select * from routestn where routeid=" & rstFare("routeid") & " order by routeStnNo asc", railCn
        rstRoute.MoveFirst
        sid = rstRoute("stnid")
        rstRoute.Close
        If rstFare("upTrainNo") = txtTrainNo Then
            temp = 0
            
        ElseIf rstFare("downTrainNo") = txtTrainNo Then
            temp = 1
            
        End If
        If temp = 0 Then
            dist1 = frmTrainSchedule.sumDist(sid, rstFare("routeid"), sourceID)
            dist2 = frmTrainSchedule.sumDist(sid, rstFare("routeid"), destID)
        Else
            dist1 = frmTrainSchedule.sumDist1(sid, rstFare("routeid"), sourceID)
            dist2 = frmTrainSchedule.sumDist1(sid, rstFare("routeid"), destID)
        End If
        dist = dist2 - dist1
        fare = fare * dist
        lblFare.Caption = "Train Name : " & trainName & Chr(vbKeyReturn) & "Class : " & cmbCoach.Text & Chr(vbKeyReturn) & "Journey Date : " & dateJourney.Value & Chr(vbKeyReturn) & txtSource & " - " & txtDest & Chr(vbKeyReturn) & "Fare(In Rs.) : " & fare
        Frame1.Width = 13695
        Image4.Left = 14640
    Else
        MsgBox "Train Not Run in Date " & dateJourney.Value, vbCritical
        Frame1.Width = 7815
        Image4.Left = 8760
    End If
Else
    MsgBox "Please Fill All Fields.", vbCritical
    Frame1.Width = 7815
    Image4.Left = 8760
End If
End Sub
Private Sub txtTrainNo_KeyPress(KeyAscii As Integer)
    Call validation(1, KeyAscii, txtTrainNo)
End Sub
Private Sub Form_load()
    Frame1.Width = 7815
    Image4.Left = 8760
    Set rstCoach = New ADODB.Recordset
    rstCoach.CursorLocation = adUseClient
    rstCoach.Open "select * from coach", railCn
    If rstCoach.RecordCount > 0 Then
        i = 0
        rstCoach.MoveFirst
        Do While Not rstCoach.EOF
            cmbCoach.AddItem rstCoach(1)
            cmbCoach.ItemData(i) = rstCoach(0)
            rstCoach.MoveNext
            i = i + 1
        Loop
    End If
    rstCoach.Close
    
    dateJourney.Value = Date
    List1.Visible = False
    List1.Top = Frame1.Top + txtSource.Top + txtSource.Height
    List1.Left = txtSource.Left + Frame1.Left
    List2.Visible = False
    List2.Top = Frame1.Top + txtDest.Top + txtDest.Height
    List2.Left = txtDest.Left + Frame1.Left
End Sub
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


Private Sub txtTrainNo_LostFocus()
    If txtTrainNo.Text <> "" Then
        If checkTrainNo() = False Then
            MsgBox "TrainNo is Incorrect", vbCritical
            txtTrainNo.SetFocus
            txtTrainNo.Text = ""
        End If
    End If
End Sub


