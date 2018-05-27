VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTrainScheduleUser 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Train Schedule"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   15420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Train Schedule"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   13335
      Begin VB.CommandButton cmdSchedule 
         Height          =   375
         Left            =   7920
         Picture         =   "frmTrainScheduleUser.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   2295
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
         Left            =   4680
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
      Begin MSFlexGridLib.MSFlexGrid flexGridStn 
         Height          =   3975
         Left            =   840
         TabIndex        =   3
         Top             =   1800
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7011
         _Version        =   393216
         BackColor       =   -2147483634
         ForeColor       =   192
         BackColorSel    =   192
         GridColor       =   255
         GridColorFixed  =   192
         GridLinesFixed  =   1
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
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Train Number:"
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
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Station List"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   11655
      End
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   360
      Picture         =   "frmTrainScheduleUser.frx":319E
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Train Schedule"
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
      Left            =   960
      TabIndex        =   6
      Top             =   0
      Width           =   3735
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmTrainScheduleUser.frx":3986
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   11400
      Picture         =   "frmTrainScheduleUser.frx":3E60
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image5 
      Height          =   5985
      Left            =   240
      Picture         =   "frmTrainScheduleUser.frx":433A
      Top             =   1440
      Width           =   825
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   14400
      Picture         =   "frmTrainScheduleUser.frx":14954
      Top             =   1440
      Width           =   825
   End
End
Attribute VB_Name = "frmTrainScheduleUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstRouteStn As ADODB.Recordset
Dim rstSumDist As ADODB.Recordset
Dim rstdays As ADODB.Recordset
Dim rstTrainNo As ADODB.Recordset
Dim f As Integer 'see distance
Dim flag As Integer 'see distance
Dim flag1 As Integer
Public Function sumDist(sourceID As Integer, routeId As Integer, stnId As Integer) As Integer
    Dim flag1 As Integer
    Set rstSumDist = New ADODB.Recordset
    rstSumDist.CursorLocation = adUseClient
    rstSumDist.Open "select stnId,distance from routestn where routeid=" & routeId & " order by routestnno asc", railCn
    rstSumDist.MoveFirst
    If sourceID <> rstSumDist(0) Then
         For j = 0 To rstSumDist.RecordCount - 1
             If sourceID = rstSumDist(0) Then
                 Exit For
             Else
                 rstSumDist.MoveNext
             End If
         Next
        
         For i = j To rstSumDist.RecordCount - 1
             rstSumDist.MoveNext
             If sourceID = stnId Then
                 flag = 0
                 Exit For
             ElseIf rstSumDist(0) <> stnId And (f <> 1 Or flag <> 0) Then
                 If f <> 1 Then
                    rstSumDist.MoveNext
                    f = 0
                    flag = flag + rstSumDist(1)
                 End If
             Else
                 f = 1
                 flag = flag + rstSumDist(1)
                 Exit For
             End If
             
             
         Next
        sumDist = flag
    Else
        For i = 0 To rstSumDist.RecordCount - 1
             If rstSumDist(0) <> stnId Then
                    flag1 = flag1 + rstSumDist(1)
             Else
                 flag1 = flag1 + rstSumDist(1)
                 Exit For
             End If
             rstSumDist.MoveNext
         Next
         sumDist = flag1
    End If
    
End Function
Public Function sumDist1(sourceID As Integer, routeId As Integer, stnId As Integer) As Integer
    
    Set rstSumDist = New ADODB.Recordset
    rstSumDist.CursorLocation = adUseClient
    rstSumDist.Open "select stnId,distance from routestn where routeid=" & routeId & " order by routestnno desc", railCn
    rstSumDist.MoveFirst
    If sourceID <> rstSumDist(0) Then
        For j = 0 To rstSumDist.RecordCount - 1
            If stnId = rstSumDist(0) Then
                Exit For
            Else
                rstSumDist.MoveNext
            End If
        Next
        
        For i = j To rstSumDist.RecordCount - 1
             If stnId = sourceID Then
                flag = 0
                Exit For
             ElseIf rstSumDist(0) <> stnId And flag <> 0 Then
                rstSumDist.MovePrevious
                If f <> 0 Then
                    rstSumDist.MoveNext
                    flag = flag + rstSumDist(1)
                    Exit For
                    
                End If
             Else
                    rstSumDist.MovePrevious
                    f = 1
                    flag = flag + rstSumDist(1)
                    Exit For
             End If
        Next
        sumDist1 = flag
    Else
        
        For j = 0 To rstSumDist.RecordCount - 1
            If stnId = rstSumDist(0) Then
                Exit For
            Else
                rstSumDist.MoveNext
            End If
        Next
        
        For i = 0 To rstSumDist.RecordCount - 1
             If stnId = sourceID Then
                flag1 = 0
                Exit For
'             ElseIf rstSumDist(0) <> stnId And i > 2 Then
''                rstSumDist.MovePrevious
''                rstSumDist.MoveNext
'                flag1 = flag1 + rstSumDist(1)
                
             Else
'                If rstSumDist(0) = stnId Then
                    rstSumDist.MovePrevious
                    flag1 = flag1 + rstSumDist(1)
                    Exit For
'                Else
'                    rstSumDist.MovePrevious
'                    flag1 = flag1 + rstSumDist(1)
'                    Exit For
'                End If
             End If

         Next
         sumDist1 = flag1
    End If
    
End Function
Public Sub cmdSchedule_Click()
If txtTrainNo.Text <> "" Then
Dim sourceID As Integer
    Set rstRouteStn = New ADODB.Recordset
    rstRouteStn.CursorLocation = adUseClient
    trainNo = txtTrainNo.Text
    rstRouteStn.Open "select  station.stncode,station.stnname,train.routeid,arrtime,deptime,day,trainname,trainroute.stnId from trainroute,station,train where( train.uptrainno=trainroute.trainno or  train.downtrainno=trainroute.trainno )and trainroute.trainno='" & trainNo & "' and station.stnid=trainroute.stnid order by stopageno asc", railCn
    If rstRouteStn.RecordCount > 0 Then
        Set rstTrainNo = New ADODB.Recordset
        rstTrainNo.CursorLocation = adUseClient
        rstTrainNo.Open "select uptrainno,downtrainno from train where uptrainno='" & trainNo & "' or downtrainno='" & trainNo & "'", railCn
        Dim temp As Integer
         rstTrainNo.MoveFirst
        For i = 0 To rstTrainNo.RecordCount - 1
            If rstTrainNo("upTrainNo") = trainNo Then
                temp = 0
                Exit For
            ElseIf rstTrainNo("downTrainNo") = trainNo Then
                temp = 1
                Exit For
            End If
            rstTrainNo.MoveNext
        Next
        flexGridStn.Rows = 1
        flexGridStn.Cols = 8
        rstRouteStn.MoveFirst
        sourceID = rstRouteStn(7)
        For i = 1 To rstRouteStn.RecordCount
            flexGridStn.Rows = flexGridStn.Rows + 1
            flexGridStn.TextMatrix(i, 0) = i
            flexGridStn.TextMatrix(i, 1) = rstRouteStn(0)
            flexGridStn.TextMatrix(i, 2) = rstRouteStn(1)
            flexGridStn.TextMatrix(i, 3) = rstRouteStn(2)
            flexGridStn.TextMatrix(i, 4) = convertTime(rstRouteStn(3))
            flexGridStn.TextMatrix(i, 5) = convertTime(rstRouteStn(4))
            flexGridStn.TextMatrix(i, 6) = rstRouteStn(5)
            If temp = 0 Then
                flexGridStn.TextMatrix(i, 7) = sumDist(sourceID, rstRouteStn(2), rstRouteStn(7))
            Else
                flexGridStn.TextMatrix(i, 7) = sumDist1(sourceID, rstRouteStn(2), rstRouteStn(7))
            End If
            rstRouteStn.MoveNext
        Next
    rstRouteStn.MoveFirst
    Label4.Caption = " Train Schedule of " & rstRouteStn("trainname") & "(" & txtTrainNo.Text & ")"
    Set rstdays = New ADODB.Recordset
    rstdays.CursorLocation = adUseClient
    rstdays.Open "select *from days where trainno='" & txtTrainNo.Text & "' ", railCn
    rstdays.MoveFirst
    For j = 1 To 7
        If rstdays(j) = True Then
            
            If Label4.Caption = " Train Schedule of " & rstRouteStn("trainname") & "(" & txtTrainNo.Text & ")" Then
                Label4.Caption = Label4.Caption & Chr(vbKeyReturn) & " Runs On |" & rstdays(j).Name & "|"
            Else
                Label4.Caption = Label4.Caption & rstdays(j).Name & "|"
            End If
        End If
    Next
    
    Else
        MsgBox "Train Not Found", vbCritical
    End If
Else
    MsgBox "Please Enter Train Number", vbCritical
End If
End Sub

Private Sub Form_load()
    f = 0
    flag = 0
    flag1 = 0
    flexGridStn.Rows = 1
    flexGridStn.Cols = 8
    flexGridStn.ColWidth(0) = 600
    flexGridStn.ColWidth(1) = 1000
    flexGridStn.ColWidth(2) = 2900
    flexGridStn.ColWidth(3) = 1000
    flexGridStn.ColWidth(4) = 2000
    flexGridStn.ColWidth(5) = 2000
    flexGridStn.ColWidth(7) = 1100
    flexGridStn.ColWidth(6) = 950
    flexGridStn.TextMatrix(0, 0) = "S.No."
    flexGridStn.TextMatrix(0, 1) = "Stn Code"
    flexGridStn.TextMatrix(0, 2) = "Stn Name"
    flexGridStn.TextMatrix(0, 3) = "Route ID"
    flexGridStn.TextMatrix(0, 4) = "Arr. Time"
    flexGridStn.TextMatrix(0, 5) = "Dep. Time"
    flexGridStn.TextMatrix(0, 6) = "Day"
    flexGridStn.TextMatrix(0, 7) = "Distance"
End Sub


Private Sub txtTrainNo_KeyPress(KeyAscii As Integer)
    Call validation(1, KeyAscii, txtTrainNo)
End Sub

