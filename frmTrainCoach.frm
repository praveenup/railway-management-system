VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTrainCoach 
   BackColor       =   &H8000000E&
   Caption         =   "Train Coach"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   14505
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
      Height          =   7455
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   14175
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   10920
         Picture         =   "frmTrainCoach.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "Search Train"
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
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   5415
         Begin VB.TextBox txtTrainName 
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
            Left            =   1920
            TabIndex        =   10
            Top             =   840
            Width           =   2895
         End
         Begin MSDataListLib.DataList DataList 
            Height          =   4155
            Left            =   960
            TabIndex        =   11
            Top             =   1680
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   7329
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bodoni MT"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label8 
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
            Left            =   480
            TabIndex        =   12
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.TextBox txtCoaches 
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
         Left            =   9240
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeleteRow 
         Height          =   375
         Left            =   13320
         Picture         =   "frmTrainCoach.frx":1902
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2760
         Width           =   375
      End
      Begin VB.ComboBox cmbCoachType 
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
         Left            =   9240
         TabIndex        =   2
         Top             =   1080
         Width           =   2895
      End
      Begin MSFlexGridLib.MSFlexGrid flexGridCoaches 
         Height          =   4215
         Left            =   7320
         TabIndex        =   4
         ToolTipText     =   "Select Field to Update or Delete Coaches"
         Top             =   2760
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7435
         _Version        =   393216
         BackColor       =   -2147483634
         ForeColor       =   192
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "COACHES TABLE"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   7320
         TabIndex        =   13
         Top             =   2280
         Width           =   6015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6960
         TabIndex        =   8
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "No. Of Coach:"
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
         Left            =   7800
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         Left            =   7800
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Image Image5 
      Height          =   5985
      Left            =   240
      Picture         =   "frmTrainCoach.frx":20B0
      Top             =   1560
      Width           =   825
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   15240
      Picture         =   "frmTrainCoach.frx":126CA
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD Train coach"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   840
      Picture         =   "frmTrainCoach.frx":22CE4
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmTrainCoach.frx":2362F
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11400
      Picture         =   "frmTrainCoach.frx":23B09
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmTrainCoach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTrain As ADODB.Recordset
Dim rstCoachType As ADODB.Recordset
Dim rstCoaches As ADODB.Recordset
Dim downTrainNo As Long, upTrainNo As Long
Dim coachinitial As String * 1
Dim cmdCoach As ADODB.Command
Dim temp As Variant
Private Sub cmbCoachType_Click()
    Set rstCoachType = New ADODB.Recordset
    rstCoachType.CursorLocation = adUseClient
    If cmbCoachType.ListIndex <> -1 Then
    rstCoachType.Open "select coachinitial from Coach where coachtypeid=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & "", railCn
    If rstCoachType.RecordCount > 0 Then
        coachinitial = rstCoachType(0)
    End If
    End If
End Sub
Private Function checkAlready() As Boolean
    If flexGridCoaches.Rows > 1 Then
        For i = 1 To flexGridCoaches.Rows - 1
            If flexGridCoaches.TextMatrix(i, 1) = cmbCoachType.ItemData(cmbCoachType.ListIndex) Then
                checkAlready = False
                Exit Function
            Else
                checkAlready = True
            End If
        Next
    End If
    checkAlready = True
End Function

Private Sub cmdAdd_Click()
If upTrainNo <> 0 Then
    If cmbCoachType.ListIndex <> -1 And txtCoaches.Text <> "" Then
        If checkAlready() Then
            Set cmdCoach = New ADODB.Command
            cmdCoach.CommandType = adCmdText
            cmdCoach.ActiveConnection = railCn
            For i = 1 To Val(txtCoaches.Text)
                temp = coachinitial & i
                cmdCoach.CommandText = "insert into traincoach values('" & upTrainNo & "'," & cmbCoachType.ItemData(cmbCoachType.ListIndex) & ",'" & temp & "')"
                cmdCoach.Execute
            Next
            For i = 1 To Val(txtCoaches.Text)
                temp = coachinitial & i
                cmdCoach.CommandText = "insert into traincoach values('" & downTrainNo & "'," & cmbCoachType.ItemData(cmbCoachType.ListIndex) & ",'" & temp & "')"
                cmdCoach.Execute
                
                flexGridCoaches.Rows = flexGridCoaches.Rows + 1
                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 0) = flexGridCoaches.Rows - 1
                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 1) = cmbCoachType.ItemData(cmbCoachType.ListIndex)
                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 2) = cmbCoachType
                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 3) = temp
            Next
            Label3.Caption = "Record Successfully Added!!!"
        Else
            Label3.Caption = "Coach Already Present,It Cannot Be Added!!!"
        End If
    Else
        Label3.Caption = "Fill All Fields!!!"
    End If
Else
    Label3.Caption = "Please Search The Train First!!!"
End If
End Sub

Private Sub cmdDeleteRow_Click()
        If flexGridCoaches.Row <> 0 And flexGridCoaches.Rows = 2 Then
            Set cmdCoach = New ADODB.Command
            cmdCoach.CommandType = adCmdText
            cmdCoach.ActiveConnection = railCn
            cmdCoach.CommandText = "delete from traincoach where (trainno='" & upTrainNo & "' or trainno='" & downTrainNo & "') and coachtypeid=" & flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) & " "
            cmdCoach.Execute
            flexGridCoaches.Rows = flexGridCoaches.Rows - 1
        Else
            Set cmdCoach = New ADODB.Command
            cmdCoach.CommandType = adCmdText
            cmdCoach.ActiveConnection = railCn
            cmdCoach.CommandText = "delete from traincoach where (trainno='" & upTrainNo & "' or trainno='" & downTrainNo & "') and coachtypeid=" & flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) & " "
            cmdCoach.Execute
    
            flexGridCoaches.Rows = 1
            flexGridCoaches.Cols = 4
            Set rstCoaches = New ADODB.Recordset
            rstCoaches.CursorLocation = adUseClient
            rstCoaches.Open "select traincoach.coachtypeid,coachtypename,coachname from traincoach,coach where coach.coachtypeid=traincoach.coachtypeid and trainno='" & dataList.BoundText & "' ", railCn
                If rstCoaches.RecordCount > 0 Then
                    rstCoaches.MoveFirst
                    For i = 0 To rstCoaches.RecordCount - 1
                        flexGridCoaches.Rows = flexGridCoaches.Rows + 1
                        flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 0) = flexGridCoaches.Rows - 1
                        flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 1) = rstCoaches(0)
                        flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 2) = rstCoaches(1)
                        flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 3) = rstCoaches(2)
                        rstCoaches.MoveNext
                    Next
                End If
        End If
        cmdDeleteRow.Visible = False
        cmdAdd.Visible = True
End Sub
'Private Function checkAlready1() As Boolean
'    If flexGridCoaches.Rows > 1 Then
'        For i = 1 To flexGridCoaches.Rows - 1
'            If i = flexGridCoaches.Row Then
'                If flexGridCoaches.TextMatrix(i, 1) = cmbCoachType.ItemData(cmbCoachType.ListIndex) And flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) <> cmbCoachType.ItemData(cmbCoachType.ListIndex) Then
'                    checkAlready1 = False
'                    Exit For
'                Else
'                    checkAlready1 = True
'                End If
'            Else
'                If flexGridCoaches.TextMatrix(i, 1) = cmbCoachType.ItemData(cmbCoachType.ListIndex) Then
'                    checkAlready1 = False
'                    Exit For
'                Else
'                    checkAlready1 = True
'                End If
'            End If
'        Next
'    End If
'End Function
'Private Sub cmdeditSave_Click()
'    If cmbCoachType.ListIndex <> -1 And txtCoaches.Text <> "" Then
''        If checkAlready1() Then
'            Set cmdCoach = New ADODB.Command
'            cmdCoach.CommandType = adCmdText
'            cmdCoach.ActiveConnection = railCn
'            cmdCoach.CommandText = "delete from traincoach where (trainno='" & upTrainNo & "' or trainno='" & downTrainNo & "') and coachtypeid=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & ""
'            cmdCoach.Execute
''            Dim j As Integer
''            j = flexGridCoaches.Rows
''            For i = 1 To j - 1
''                If flexGridCoaches.TextMatrix(i, 1) = cmbCoachType.ItemData(cmbCoachType.ListIndex) Then
''                    flexGridCoaches.RemoveItem (i)
''                    j = j - 1
''                End If
''            Next
'            For i = 1 To Val(txtCoaches.Text)
'                temp = coachinitial & i
'                cmdCoach.CommandText = "insert into traincoach values('" & upTrainNo & "'," & cmbCoachType.ItemData(cmbCoachType.ListIndex) & ",'" & temp & "')"
'                cmdCoach.Execute
'            Next
'            For i = 1 To Val(txtCoaches.Text)
'                temp = coachinitial & i
'                cmdCoach.CommandText = "insert into traincoach values('" & downTrainNo & "'," & cmbCoachType.ItemData(cmbCoachType.ListIndex) & ",'" & temp & "')"
'                cmdCoach.Execute
'
'                flexGridCoaches.Rows = flexGridCoaches.Rows + 1
'                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 0) = flexGridCoaches.Rows - 1
'                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 1) = cmbCoachType.ItemData(cmbCoachType.ListIndex)
'                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 2) = cmbCoachType
'                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 3) = temp
'            Next
'            Label3.Caption = "Record Successfully Updated!!!"
''        Else
''            Label3.Caption = "Coach Already Present,It Cannot Be Added!!!"
''        End If
'    Else
'        Label3.Caption = "Fill All Fields!!!"
'    End If
'    cmdAdd.Visible = True
'    cmdeditSave.Visible = False
'    cmdDeleteRow.Visible = False
'End Sub

Private Sub DataList_Click()
    If flexGridCoaches.Rows = 1 Or temp <> dataList.Text Then
        flexGridCoaches.Rows = 1
        flexGridCoaches.Cols = 4
        Set rstTrain = New ADODB.Recordset
        rstTrain.CursorLocation = adUseClient
        rstTrain.Open "select downtrainno,trainname from train where uptrainno='" & dataList.BoundText & "' ", railCn
        downTrainNo = rstTrain(0)
        upTrainNo = dataList.BoundText
        temp = rstTrain(1)
            Set rstCoaches = New ADODB.Recordset
            rstCoaches.CursorLocation = adUseClient
            rstCoaches.Open "select traincoach.coachtypeid,coachtypename,coachname from traincoach,coach where coach.coachtypeid=traincoach.coachtypeid and trainno='" & dataList.BoundText & "' ", railCn
                If rstCoaches.RecordCount > 0 Then
                    rstCoaches.MoveFirst
                    For i = 0 To rstCoaches.RecordCount - 1
                        flexGridCoaches.Rows = flexGridCoaches.Rows + 1
                        flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 0) = flexGridCoaches.Rows - 1
                        flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 1) = rstCoaches(0)
                        flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 2) = rstCoaches(1)
                        flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 3) = rstCoaches(2)
                        rstCoaches.MoveNext
                    Next
                Else
                    flexGridCoaches.Rows = 1
                    flexGridCoaches.Cols = 4
                End If
            Label3.Caption = "Fill Below Required Fields to Add Coach!!!"
            cmbCoachType.ListIndex = -1
            txtCoaches.Text = ""
            Label4.Caption = "Coaches of Train " & dataList.Text & "(" & upTrainNo & "," & downTrainNo & ")"
    End If
End Sub
Private Function totalCoach() As Integer
    Dim sum As Integer
    For i = 1 To flexGridCoaches.Rows - 1
        If flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) = flexGridCoaches.TextMatrix(i, 1) Then
            sum = sum + 1
        End If
    Next
    totalCoach = sum
End Function
Private Sub flexGridCoaches_Click()
    If flexGridCoaches.Row > 0 Then
        For i = 0 To cmbCoachType.ListCount - 1
            If cmbCoachType.ItemData(i) = flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) Then
                cmbCoachType.ListIndex = i
            End If
        Next
        txtCoaches.Text = totalCoach()
        cmdAdd.Visible = False
               cmdDeleteRow.Visible = True
        cmdDeleteRow.Top = flexGridCoaches.CellTop + flexGridCoaches.Top
    End If
End Sub

Private Sub Form_load()
    Label3.Caption = ""
    Set rstTrain = New ADODB.Recordset
    rstTrain.CursorLocation = adUseClient
    rstTrain.Open "select uptrainno,downtrainno,trainname from train order by trainname ", railCn
    If rstTrain.RecordCount > 0 Then
        Set dataList.RowSource = rstTrain
        Set dataList.DataSource = rstTrain
        dataList.BoundColumn = rstTrain.Fields(0).Name
        dataList.ListField = rstTrain.Fields(2).Name
    End If
    Set rstCoachType = New ADODB.Recordset
    rstCoachType.CursorLocation = adUseClient
    rstCoachType.Open "select * from Coach", railCn
    If rstCoachType.RecordCount > 0 Then
        i = 0
        rstCoachType.MoveFirst
        Do While Not rstCoachType.EOF
            cmbCoachType.AddItem rstCoachType(1)
            cmbCoachType.ItemData(i) = rstCoachType(0)
            rstCoachType.MoveNext
            i = i + 1
        Loop
    End If
    rstCoachType.Close
    flexGridCoaches.Rows = 1
    flexGridCoaches.Cols = 4
    flexGridCoaches.TextMatrix(0, 0) = "S.No."
    flexGridCoaches.TextMatrix(0, 1) = "CoachID"
    flexGridCoaches.TextMatrix(0, 2) = "Coach Type Name"
    flexGridCoaches.TextMatrix(0, 3) = "Coach Name"
    flexGridCoaches.ColWidth(0) = 600
    flexGridCoaches.ColWidth(1) = 1000
    flexGridCoaches.ColWidth(2) = 3000
    flexGridCoaches.ColWidth(3) = 1300
    cmdDeleteRow.Visible = False
    End Sub

Private Sub Picture1_Click()

End Sub

Private Sub txtTrainName_Change()
    Set rstTrain = New ADODB.Recordset
    rstTrain.CursorLocation = adUseClient
    rstTrain.Open "select uptrainno,downtrainno,trainname from train where trainname like '%" & txtTrainName & "%' order by trainname ", railCn
    If rstTrain.RecordCount > 0 Then
        Set dataList.RowSource = rstTrain
        Set dataList.DataSource = rstTrain
        dataList.BoundColumn = rstTrain.Fields(0).Name
        dataList.ListField = rstTrain.Fields(2).Name
    End If
End Sub
