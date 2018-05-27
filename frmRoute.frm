VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRoute 
   BackColor       =   &H8000000E&
   Caption         =   "Route"
   ClientHeight    =   10395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   17970
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   120
      Top             =   3600
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   8880
      Picture         =   "frmRoute.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   7080
      Picture         =   "frmRoute.frx":28FE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   10680
      Picture         =   "frmRoute.frx":5178
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   5280
      Picture         =   "frmRoute.frx":7A76
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7920
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
      Height          =   7095
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   15255
      Begin VB.TextBox txtDestDistance 
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
         Left            =   11160
         TabIndex        =   23
         ToolTipText     =   "other"
         Top             =   6480
         Width           =   2895
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   9480
         Picture         =   "frmRoute.frx":A1B8
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbDest 
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
         Left            =   4680
         TabIndex        =   11
         Top             =   6480
         Width           =   2895
      End
      Begin VB.ComboBox cmbSource 
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
         ItemData        =   "frmRoute.frx":BABA
         Left            =   6360
         List            =   "frmRoute.frx":BABC
         TabIndex        =   10
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "Add Intermediate Stations"
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
         Height          =   4215
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   14775
         Begin VB.CommandButton cmdEditSave 
            Height          =   375
            Left            =   6600
            Picture         =   "frmRoute.frx":BABE
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton cmdDeleteRow 
            Height          =   375
            Left            =   14280
            Picture         =   "frmRoute.frx":D4C8
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtStopage 
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
            Left            =   3480
            TabIndex        =   19
            ToolTipText     =   "other"
            Top             =   2760
            Width           =   2895
         End
         Begin VB.ComboBox cmbInterStn 
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
            Left            =   3480
            TabIndex        =   9
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txtDistance 
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
            Left            =   3480
            TabIndex        =   8
            Top             =   1920
            Width           =   2895
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   375
            Left            =   6600
            Picture         =   "frmRoute.frx":DC76
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid flexGridRoute 
            Height          =   3615
            Left            =   8160
            TabIndex        =   26
            ToolTipText     =   "Select Field to Update or Delete Coaches"
            Top             =   480
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   6376
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
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            Caption         =   "Route Station List"
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
            Height          =   255
            Left            =   8160
            TabIndex        =   27
            Top             =   240
            Width           =   6135
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000E&
            Caption         =   "Stopage Number:  (Other than 1)"
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
            Left            =   840
            TabIndex        =   18
            Top             =   2640
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            Caption         =   "Distance Travel:     (From Previous Station)"
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
            Left            =   840
            TabIndex        =   7
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            Caption         =   "Station Name:"
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
            Left            =   840
            TabIndex        =   4
            Top             =   1080
            Width           =   1935
         End
      End
      Begin VB.TextBox txtRoute 
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
         Left            =   6360
         TabIndex        =   2
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Destination Station Distance:  (from last intermediate stn)"
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
         Left            =   8040
         TabIndex        =   22
         Top             =   6360
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Destination Station:"
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
         Left            =   2640
         TabIndex        =   12
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Source Station:"
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
         Left            =   4320
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   4320
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   16200
      Picture         =   "frmRoute.frx":F578
      Top             =   1560
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   5985
      Left            =   120
      Picture         =   "frmRoute.frx":1FB92
      Top             =   1560
      Width           =   825
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
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
      Height          =   255
      Left            =   3360
      TabIndex        =   25
      Top             =   480
      Width           =   12615
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11520
      Picture         =   "frmRoute.frx":301AC
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   1200
      Picture         =   "frmRoute.frx":30686
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD ROUTE INFORMATION"
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
      Left            =   1920
      TabIndex        =   24
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmRoute.frx":30EE2
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmRoute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdRoute As ADODB.Command
Dim rstRoute As ADODB.Recordset
Dim rstRoute1 As ADODB.Recordset
Dim rstStn As ADODB.Recordset
Dim ID As Integer
Dim check As Integer 'for intermediate stn is already present or not
Dim max As Integer 'for dest stn stopage no.
Dim missingStopage As Integer
Private Function checkStopSequence() As Integer
    If flexGridRoute.Rows > 1 Then
        Dim s() As Integer
        ReDim s(flexGridRoute.Rows - 1) As Integer
        For i = 1 To flexGridRoute.Rows - 1
            s(i - 1) = flexGridRoute.TextMatrix(i, 2)
        Next
        Do
            For i = 2 To flexGridRoute.Rows
                For j = 1 To flexGridRoute.Rows
                    If s(j - 1) = i Then
                        checkStopSequence = 1
                        Exit For
                    Else
                        If j = flexGridRoute.Rows Then
                            checkStopSequence = 0
                            missingStopage = i
                            Exit Do
                        End If
                    End If
                Next
            Next
        Loop While False
    Else
        checkStopSequence = 1
    End If
End Function

Private Function checkStn() As Integer
    If flexGridRoute.Rows > 1 Then
        For i = 1 To flexGridRoute.Rows - 1
            If flexGridRoute.TextMatrix(i, 1) = cmbInterStn.Text Then
                checkStn = 1
                Exit For
            Else
                checkStn = 0
            End If
        Next
    End If
End Function
Private Function checkStn1() As Integer
    If flexGridRoute.Rows > 1 Then
        For i = 1 To flexGridRoute.Rows - 1
            If i = flexGridRoute.Row Then
                If flexGridRoute.TextMatrix(i, 1) = cmbInterStn.Text And flexGridRoute.TextMatrix(flexGridRoute.Row, 1) <> cmbInterStn.Text Then
                    checkStn1 = 1
                    Exit For
                Else
                    checkStn1 = 0
                End If
            Else
                If flexGridRoute.TextMatrix(i, 1) = cmbInterStn.Text Then
                    checkStn1 = 1
                    Exit For
                Else
                    checkStn1 = 0
                End If
            End If
        Next
    End If
End Function
Private Function checkStopage() As Integer
    If flexGridRoute.Rows > 1 Then
        For i = 1 To flexGridRoute.Rows - 1
            If Val(txtStopage.Text) = flexGridRoute.TextMatrix(i, 2) Then
                checkStopage = 1
                Exit For
            Else
                checkStopage = 0
            End If
        Next
    End If
End Function
Private Function checkStopage1() As Integer
    If flexGridRoute.Rows > 1 Then
        For i = 1 To flexGridRoute.Rows - 1
            If i = flexGridRoute.Row Then
                If Val(txtStopage.Text) = flexGridRoute.TextMatrix(i, 2) And flexGridRoute.TextMatrix(flexGridRoute.Row, 2) <> Val(txtStopage.Text) Then
                    checkStopage1 = 1
                    Exit For
                Else
                    checkStopage1 = 0
                End If
            Else
                If Val(txtStopage.Text) = flexGridRoute.TextMatrix(i, 2) Then
                    checkStopage1 = 1
                    Exit For
                Else
                    checkStopage1 = 0
                End If
            End If
        Next
    End If
End Function

Private Sub cmbSource_Click()
    flexGridRoute.Rows = 1
    flexGridRoute.Cols = 4
    flexGridRoute.FixedCols = 1
    cmbInterStn.Text = ""
    txtDistance.Text = ""
    txtStopage.Text = ""
    cmdAdd.Visible = True
    cmdeditSave.Visible = False
    cmdDeleteRow.Visible = False
End Sub



Private Sub cmdAdd_Click()
    If cmbSource.ListIndex <> -1 And cmbInterStn.ListIndex <> -1 And txtStopage <> "" And txtDistance <> "" And txtRoute <> "" Then
            If (cmbInterStn.ListIndex <> cmbSource.ListIndex) Then
                check = checkStn()
                If check = 0 Then
                        check = checkStopage()
                        If check = 0 Then
                                If txtStopage <> 1 Then
                                    flexGridRoute.Rows = flexGridRoute.Rows + 1
                                    flexGridRoute.TextMatrix(flexGridRoute.Rows - 1, 0) = cmbInterStn.ItemData(cmbInterStn.ListIndex)
                                    flexGridRoute.TextMatrix(flexGridRoute.Rows - 1, 2) = txtStopage
                                    flexGridRoute.TextMatrix(flexGridRoute.Rows - 1, 1) = cmbInterStn
                                    flexGridRoute.TextMatrix(flexGridRoute.Rows - 1, 3) = txtDistance
                                    cmbInterStn.Text = ""
                                    txtDistance.Text = ""
                                    txtStopage.Text = ""
                                Else
                                    MsgBox "Stopage Number 1 Already Alloted to Source Station", vbCritical
                                End If
                        Else
                            MsgBox "Stopage Number is Already Added,Give Another Stopage Number", vbCritical
                        End If
                Else
                    MsgBox "Intermediate Station Already Added", vbCritical
                End If
            Else
                MsgBox "Intermediate Station must be Different from Source Station", vbCritical
            End If
    Else
        MsgBox "Please Fill all Fields", vbCritical
    End If
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
    If saveUpdate = 2 Then
        If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
                Set cmdRoute = New ADODB.Command
                cmdRoute.CommandType = adCmdText
                cmdRoute.ActiveConnection = railCn
                cmdRoute.CommandText = "delete from routeStn where routeID=" & frmRouteDialog.routeId & ""
                cmdRoute.Execute
                cmdRoute.CommandText = "delete from route where routeID=" & frmRouteDialog.routeId & ""
                cmdRoute.Execute
                MsgBox "Record Successfully Deleted", vbInformation
                txtRoute.Text = ""
                cmbSource.Text = ""
                cmbDest.Text = ""
                cmbInterStn.Text = ""
                txtDistance.Text = ""
                txtStopage.Text = ""
                txtDestDistance.Text = ""
                txtDestDistance.BackColor = vbButtonFace
                txtStopage.BackColor = vbButtonFace
                txtDistance.BackColor = vbButtonFace
                cmbInterStn.BackColor = vbButtonFace
                cmbDest.BackColor = vbButtonFace
                cmbSource.BackColor = vbButtonFace
                txtRoute.BackColor = vbButtonFace
                flexGridRoute.Rows = 1
                flexGridRoute.Cols = 4
                flexGridRoute.FixedCols = 1
                txtRoute.Enabled = False
                cmbSource.Enabled = False
                cmbDest.Enabled = False
                cmbInterStn.Enabled = False
                txtDistance.Enabled = False
                txtStopage.Enabled = False
                cmdeditSave.Visible = False
                cmdDeleteRow.Visible = False
                txtDestDistance.Enabled = False
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

Private Sub cmdEditSave_Click()
    If cmbSource.ListIndex <> -1 And cmbInterStn.ListIndex <> -1 And txtStopage <> "" And txtDistance <> "" And txtRoute <> "" Then
            If (cmbInterStn.ListIndex <> cmbSource.ListIndex) Then
                check = checkStn1()
                If check = 0 Then
                        check = checkStopage1()
                        If check = 0 Then
                                If txtStopage <> 1 Then
                                    flexGridRoute.TextMatrix(flexGridRoute.Row, 0) = cmbInterStn.ItemData(cmbInterStn.ListIndex)
                                    flexGridRoute.TextMatrix(flexGridRoute.Row, 2) = txtStopage
                                    flexGridRoute.TextMatrix(flexGridRoute.Row, 1) = cmbInterStn
                                    flexGridRoute.TextMatrix(flexGridRoute.Row, 3) = txtDistance
                                    cmdAdd.Visible = True
                                    cmdeditSave.Visible = False
                                    cmdDeleteRow.Visible = False
                                    cmbInterStn.Text = ""
                                    txtDistance.Text = ""
                                    txtStopage.Text = ""
                                Else
                                    MsgBox "Stopage Number 1 Already Alloted to Source Station", vbCritical
                                End If
                            
                        Else
                            MsgBox "Stopage Number is Already Added,Give Another Stopage Number", vbCritical
                        End If
                Else
                    MsgBox "Intermediate Station Already Added", vbCritical
                End If
            Else
                MsgBox "Intermediate Station must be Different from Source Station", vbCritical
            End If
    Else
        MsgBox "Please Fill all Fields", vbCritical
    End If
End Sub

Private Sub cmdNew_Click()
    txtRoute.Enabled = True
    cmbSource.Enabled = True
    cmbDest.Enabled = True
    cmbInterStn.Enabled = True
    txtDistance.Enabled = True
    txtStopage.Enabled = True
    txtDestDistance.Enabled = True
    txtRoute.BackColor = vbHighlightText
    cmbSource.BackColor = vbHighlightText
    cmbDest.BackColor = vbHighlightText
    cmbInterStn.BackColor = vbHighlightText
    txtDistance.BackColor = vbHighlightText
    txtStopage.BackColor = vbHighlightText
    txtDestDistance.BackColor = vbHighlightText
    saveUpdate = 1
    txtRoute.Text = ""
    cmbSource.Text = ""
    cmbDest.Text = ""
    cmbInterStn.Text = ""
    txtDistance.Text = ""
    txtStopage.Text = ""
    txtDestDistance.Text = ""
    flexGridRoute.Rows = 1
    flexGridRoute.Cols = 4
    flexGridRoute.FixedCols = 1
End Sub

Private Function checkInter() As Integer 'for checking dest stn in grid is already present or not
    For i = 1 To flexGridRoute.Rows - 1
        If flexGridRoute.TextMatrix(i, 0) = cmbDest.ItemData(cmbDest.ListIndex) Then
            checkInter = 1
            Exit For
        Else
            checkInter = 0
        End If
    Next
End Function

Private Function maxStopage() As Integer
    Dim temp As Integer
    If flexGridRoute.Row > 0 Then
        temp = flexGridRoute.TextMatrix(1, 2)
        If flexGridRoute.Rows > 2 Then
            For i = 2 To flexGridRoute.Rows - 1
                If temp < flexGridRoute.TextMatrix(i, 2) Then
                    temp = flexGridRoute.TextMatrix(i, 2)
                End If
            Next
            maxStopage = temp
        Else
            maxStopage = temp
        End If
    Else
        maxStopage = 1
    End If
End Function
Private Function checkAlready() As Boolean
    Set rstRoute1 = New ADODB.Recordset
    rstRoute1.CursorLocation = adUseClient
    rstRoute1.Open "select * from route ", railCn
    If rstRoute1.RecordCount > 0 Then
        rstRoute1.MoveFirst
        For i = 0 To rstRoute1.RecordCount - 1
            If rstRoute1("routeName") = txtRoute.Text Then
                checkAlready = False
                rstRoute1.Close
                Exit Function
            End If
            rstRoute1.MoveNext
        Next
    End If
    checkAlready = True
    rstRoute1.Close
End Function
Private Function checkAlready1() As Boolean
    Set rstRoute1 = New ADODB.Recordset
    rstRoute1.CursorLocation = adUseClient
    rstRoute1.Open "select * from route ", railCn
    If rstRoute1.RecordCount > 0 Then
        rstRoute1.MoveFirst
        For i = 0 To rstRoute1.RecordCount - 1
            If rstRoute1("routeName") = txtRoute.Text And frmRouteDialog.routeName <> txtRoute.Text Then
                checkAlready1 = False
                rstRoute1.Close
                Exit Function
            End If
            rstRoute1.MoveNext
        Next
    End If
    checkAlready1 = True
    rstRoute1.Close
End Function
Private Sub cmdSave_Click()

If saveUpdate = 1 Or saveUpdate = 2 Then
    If cmbSource.ListIndex <> -1 And cmbDest.ListIndex <> -1 And txtDestDistance.Text <> "" And txtRoute <> "" Then
        If checkStopSequence = 1 Then
            If cmbSource.ListIndex <> cmbDest.ListIndex Then
                check = checkInter()
                If check = 0 Then
                    Set cmdRoute = New ADODB.Command
                    Set rstRoute = New ADODB.Recordset
                    cmdRoute.CommandType = adCmdText
                    cmdRoute.ActiveConnection = railCn
                    If saveUpdate = 1 Then
                        If checkAlready() Then
                            rstRoute.Open "select max(routeID) from route", railCn
                            If rstRoute.Fields(0) > 0 Then
                                ID = rstRoute.Fields(0) + 1
                            Else
                                ID = 1
                            End If
                                    
                            cmdRoute.CommandText = "insert into route values(" & ID & ",'" & txtRoute & "')"
                            cmdRoute.Execute
                            cmdRoute.CommandText = "insert into routestn values(" & ID & "," & 1 & "," & cmbSource.ItemData(cmbSource.ListIndex) & "," & 0 & ")"
                            cmdRoute.Execute
                            For i = 1 To flexGridRoute.Rows - 1
                                cmdRoute.CommandText = "insert into routestn values(" & ID & "," & flexGridRoute.TextMatrix(i, 2) & "," & flexGridRoute.TextMatrix(i, 0) & "," & flexGridRoute.TextMatrix(i, 3) & ")"
                                cmdRoute.Execute
                            Next
                            max = maxStopage() + 1
                            cmdRoute.CommandText = "insert into routestn values(" & ID & "," & max & "," & cmbDest.ItemData(cmbDest.ListIndex) & "," & txtDestDistance & ")"
                            cmdRoute.Execute
                            saveUpdate = 0
                            MsgBox "Record Successfully Saved.", vbInformation
                        Else
                            MsgBox "Route Information Already Exists, Please Give Another route Name", vbCritical
                            Exit Sub
                        End If
                    ElseIf saveUpdate = 2 Then
                        If checkAlready1() Then
                            cmdRoute.CommandText = "update route set routeName='" & txtRoute.Text & "' where routeid=" & frmRouteDialog.routeId & ""
                            cmdRoute.Execute
                            cmdRoute.CommandText = "update routeStn set stnID=" & cmbSource.ItemData(cmbSource.ListIndex) & " where routeStnNo=" & 1 & " and routeID=" & frmRouteDialog.routeId & ""
                            cmdRoute.Execute
                            cmdRoute.CommandText = "delete from routeStn where routeID=" & frmRouteDialog.routeId & " and routestnno <> " & 1 & ""
                            cmdRoute.Execute
                            For i = 1 To flexGridRoute.Rows - 1
                                cmdRoute.CommandText = "insert into routestn values(" & frmRouteDialog.routeId & "," & flexGridRoute.TextMatrix(i, 2) & "," & flexGridRoute.TextMatrix(i, 0) & "," & flexGridRoute.TextMatrix(i, 3) & ")"
                                cmdRoute.Execute
                            Next
                            max = maxStopage() + 1
                            cmdRoute.CommandText = "insert into routestn values(" & frmRouteDialog.routeId & "," & max & "," & cmbDest.ItemData(cmbDest.ListIndex) & "," & txtDestDistance & ")"
                            cmdRoute.Execute
        
                            saveUpdate = 0
                            MsgBox "Record Successfully Updated.", vbInformation
                        Else
                            MsgBox "Route Information Already Exists, Please Give Another route Name", vbCritical
                            Exit Sub
                        End If
                    End If
                    txtRoute.Text = ""
                    cmbSource.Text = ""
                    cmbDest.Text = ""
                    cmbInterStn.Text = ""
                    txtDistance.Text = ""
                    txtStopage.Text = ""
                    txtDestDistance.Text = ""
                    txtDestDistance.BackColor = vbButtonFace
                    txtStopage.BackColor = vbButtonFace
                    txtDistance.BackColor = vbButtonFace
                    cmbInterStn.BackColor = vbButtonFace
                    cmbDest.BackColor = vbButtonFace
                    cmbSource.BackColor = vbButtonFace
                    txtRoute.BackColor = vbButtonFace
                    flexGridRoute.Rows = 1
                    flexGridRoute.Cols = 4
                    flexGridRoute.FixedCols = 1
                    txtRoute.Enabled = False
                    cmbSource.Enabled = False
                    cmbDest.Enabled = False
                    cmbInterStn.Enabled = False
                    txtDistance.Enabled = False
                    txtStopage.Enabled = False
                    cmdeditSave.Visible = False
                    cmdDeleteRow.Visible = False
                    txtDestDistance.Enabled = False
                Else
                    MsgBox "Destination Station must be Different from Intermediate Station", vbCritical
                End If
            Else
                MsgBox "Source and Destination Station must be Different", vbCritical
            End If
        Else
            MsgBox "Sequencing of Station Stopage Number is not Correct, " & missingStopage & "th stopage is not Present in Stopage list", vbCritical
        End If
    Else
        MsgBox "Please Fill all Fields.", vbCritical
    End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record.", vbCritical
End If
End Sub


Private Sub cmdSearch_Click()
    frmRouteDialog.Show 1
End Sub

Private Sub cmdDeleteRow_Click()
        If flexGridRoute.Row <> 0 And flexGridRoute.Rows = 2 Then
            flexGridRoute.Rows = flexGridRoute.Rows - 1
        Else
            flexGridRoute.RemoveItem (flexGridRoute.Row)
        End If
        cmdDeleteRow.Visible = False
        cmdAdd.Visible = True
        cmdeditSave.Visible = False
 
End Sub

Private Sub flexGridRoute_Click()
    If flexGridRoute.Row > 0 Then
        For i = 0 To cmbInterStn.ListCount - 1
            If cmbInterStn.ItemData(i) = flexGridRoute.TextMatrix(flexGridRoute.Row, 0) Then
                cmbInterStn.ListIndex = i
            End If
        Next
        txtStopage.Text = flexGridRoute.TextMatrix(flexGridRoute.Row, 2)
        txtDistance.Text = flexGridRoute.TextMatrix(flexGridRoute.Row, 3)
        cmdAdd.Visible = False
        cmdeditSave.Visible = True
        cmdDeleteRow.Visible = True
        cmdDeleteRow.Top = flexGridRoute.CellTop + flexGridRoute.Top
    End If
End Sub

Private Sub Form_load()
    Set rstStn = New ADODB.Recordset
    rstStn.CursorLocation = adUseClient
    rstStn.Open "select * from station", railCn
    If rstStn.RecordCount > 0 Then
        i = 0
        rstStn.MoveFirst
        Do While Not rstStn.EOF
            cmbSource.AddItem rstStn(2)
            cmbSource.ItemData(i) = rstStn(0)
            cmbDest.AddItem rstStn(2)
            cmbDest.ItemData(i) = rstStn(0)
            cmbInterStn.AddItem rstStn(2)
            cmbInterStn.ItemData(i) = rstStn(0)
            rstStn.MoveNext
            i = i + 1
        Loop
    End If
    rstStn.Close
    Label9.Caption = " Please click Add New Button to Add New Record (OR) Search and Select the Record for Updating Existing Record."
    flexGridRoute.Rows = 1
    flexGridRoute.Cols = 4
    flexGridRoute.FixedCols = 1
    flexGridRoute.TextMatrix(0, 0) = "StnID"
    flexGridRoute.TextMatrix(0, 1) = "Station Name"
    flexGridRoute.TextMatrix(0, 2) = "Stopage No."
    flexGridRoute.TextMatrix(0, 3) = "Distance Travel"
    flexGridRoute.ColWidth(0) = 1000
    flexGridRoute.ColWidth(1) = 2000
    flexGridRoute.ColWidth(2) = 1300
    flexGridRoute.ColWidth(3) = 1730
    txtRoute.Enabled = False
    cmbSource.Enabled = False
    cmbDest.Enabled = False
    cmbInterStn.Enabled = False
    txtDistance.Enabled = False
    txtStopage.Enabled = False
    cmdeditSave.Visible = False
    cmdDeleteRow.Visible = False
    txtDestDistance.Enabled = False
End Sub





Private Sub Timer1_Timer()
    strs = Mid(Label9.Caption, 1, 1)
    Label9.Caption = Mid(Label9.Caption, 2, Len(Label9.Caption)) & strs
End Sub

Private Sub txtDestDistance_KeyPress(KeyAscii As Integer)
    Call validation(1, KeyAscii, txtDestDistance)
End Sub

Private Sub txtDistance_KeyPress(KeyAscii As Integer)
    Call validation(1, KeyAscii, txtDistance)
End Sub

Private Sub txtRoute_KeyPress(KeyAscii As Integer)
    Call validation(2, KeyAscii, txtRoute)
End Sub

Private Sub txtStopage_KeyPress(KeyAscii As Integer)
    Call validation(1, KeyAscii, txtStopage)
End Sub
