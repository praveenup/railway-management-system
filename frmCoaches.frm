VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCoaches 
   BackColor       =   &H8000000E&
   Caption         =   "Coaches"
   ClientHeight    =   8955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8955
   ScaleWidth      =   8715
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   7920
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
      Height          =   7695
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   6735
      Begin VB.CommandButton cmdSave 
         Height          =   375
         Left            =   4800
         Picture         =   "frmCoaches.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
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
         Left            =   1800
         TabIndex        =   9
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtCoachNo 
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
         Left            =   1800
         TabIndex        =   6
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton cmdDeleteRow 
         Height          =   375
         Left            =   5160
         Picture         =   "frmCoaches.frx":1A0A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdEditSave 
         Height          =   375
         Left            =   4320
         Picture         =   "frmCoaches.frx":21B8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid flexGridCoaches 
         Height          =   4215
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Select Field to Update or Delete Coaches"
         Top             =   3240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7435
         _Version        =   393216
         BackColor       =   -2147483634
         ForeColor       =   192
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
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Coach No.          (10 Digit):"
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
         TabIndex        =   7
         Top             =   1920
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
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   7680
      Picture         =   "frmCoaches.frx":3BC2
      Top             =   1680
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   5985
      Left            =   120
      Picture         =   "frmCoaches.frx":141DC
      Top             =   1680
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11280
      Picture         =   "frmCoaches.frx":247F6
      Top             =   0
      Width           =   11535
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD COACHES"
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
      TabIndex        =   2
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   600
      Picture         =   "frmCoaches.frx":24CD0
      Top             =   0
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmCoaches.frx":254B8
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmCoaches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdCoaches As ADODB.Command
Dim rstCoachType As ADODB.Recordset
Dim rstCoaches As ADODB.Recordset
Dim check As Integer 'for intermediate Coach no. is already present or not

Private Function checkInterCoach() As Integer
    If flexGridCoaches.Rows > 1 Then
        For i = 1 To flexGridCoaches.Rows - 1
            If Val(txtCoachNo.Text) = flexGridCoaches.TextMatrix(i, 1) Then
                checkInterCoach = 1
                Exit For
            Else
                checkInterCoach = 0
            End If
        Next
    End If
End Function

Private Sub cmbCoachType_Click()
   If cmbCoachType.ListIndex <> -1 Then
         flexGridCoaches.Rows = 1
         flexGridCoaches.Cols = 2
         flexGridCoaches.FixedCols = 1
        Set rstCoaches = New ADODB.Recordset
         rstCoaches.CursorLocation = adUseClient
         rstCoaches.Open "select * from Coaches where coachtypeid=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & "", railCn
         If rstCoaches.RecordCount > 0 Then
             rstCoaches.MoveFirst
             For i = 0 To rstCoaches.RecordCount - 1
                 flexGridCoaches.Rows = flexGridCoaches.Rows + 1
                 flexGridCoaches.TextMatrix(i + 1, 0) = rstCoaches(0)
                 flexGridCoaches.TextMatrix(i + 1, 1) = rstCoaches(1)
                 rstCoaches.MoveNext
             Next
         End If
        rstCoaches.Close
        txtCoachNo.Text = ""
        cmdeditSave.Visible = False
        cmdDeleteRow.Visible = False
        cmdSave.Visible = True
    End If
End Sub

Private Function checkInterCoach1() As Integer
    If flexGridCoaches.Rows > 1 Then
        For i = 1 To flexGridCoaches.Rows - 1
            If i = flexGridCoaches.Row Then
                If flexGridCoaches.TextMatrix(i, 0) = txtCoachNo.Text And flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) <> txtCoachNo.Text Then
                    checkInterCoach1 = 1
                    Exit For
                Else
                    checkInterCoach1 = 0
                End If
            Else
                If flexGridCoaches.TextMatrix(i, 0) = txtCoachNo.Text Then
                    checkInterCoach1 = 1
                    Exit For
                Else
                    checkInterCoach1 = 0
                End If
            End If
        Next
    End If
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDeleteRow_Click()
If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
    If flexGridCoaches.Row > 0 And flexGridCoaches.Rows = 2 Then
        Set cmdCoaches = New ADODB.Command
        cmdCoaches.CommandType = adCmdText
        cmdCoaches.ActiveConnection = railCn
        cmdCoaches.CommandText = "delete from coaches where CoachtypeID=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & " and coachno='" & flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) & "'"
        cmdCoaches.Execute
        flexGridCoaches.Rows = flexGridCoaches.Rows - 1
        MsgBox "Record Successfully Deleted", vbInformation
    Else
        Set cmdCoaches = New ADODB.Command
        cmdCoaches.CommandType = adCmdText
        cmdCoaches.ActiveConnection = railCn
        cmdCoaches.CommandText = "delete from coaches where CoachtypeID=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & " and coachno='" & flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) & "'"
        cmdCoaches.Execute
        flexGridCoaches.RemoveItem (flexGridCoaches.Row)
        MsgBox "Record Successfully Deleted", vbInformation
    End If
    cmdDeleteRow.Visible = False
    cmdSave.Visible = True
    cmdeditSave.Visible = False
    txtCoachNo.Text = ""
End If
End Sub

Private Sub cmdEditSave_Click()
        If cmbCoachType.ListIndex <> -1 And txtCoachNo <> "" Then
            If Len(txtCoachNo.Text) = 10 Then
                check = checkInterCoach1()
                If check = 0 Then
                        Set cmdCoaches = New ADODB.Command
                        cmdCoaches.CommandType = adCmdText
                        cmdCoaches.ActiveConnection = railCn
                        cmdCoaches.CommandText = "update Coaches set coachno=" & Val(txtCoachNo.Text) & " where coachtypeid=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & " and coachno= '" & flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) & "' "
                        cmdCoaches.Execute
                        flexGridCoaches.TextMatrix(flexGridCoaches.Row, 0) = cmbCoachType.ItemData(cmbCoachType.ListIndex)
                        flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1) = txtCoachNo.Text
                        cmdSave.Visible = True
                        cmdeditSave.Visible = False
                        cmdDeleteRow.Visible = False
                        txtCoachNo.Text = ""
                        MsgBox "Record Successfully Updated.", vbInformation
                Else
                    MsgBox "Coach No. Already Added", vbCritical
                End If
            Else
                MsgBox "Length Of The Coach No. Must Have 10 Digits", vbCritical
            End If
        Else
            MsgBox "Please Fill all Fields", vbCritical
        End If
End Sub

Private Sub cmdSave_Click()
    If cmbCoachType.ListIndex <> -1 And txtCoachNo <> "" Then
        If Len(txtCoachNo.Text) = 10 Then
            check = checkInterCoach()
            If check = 0 Then
                Set cmdCoaches = New ADODB.Command
                cmdCoaches.CommandType = adCmdText
                cmdCoaches.ActiveConnection = railCn
                cmdCoaches.CommandText = "insert into Coaches values(" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & ",'" & txtCoachNo.Text & "')"
                cmdCoaches.Execute
                flexGridCoaches.Rows = flexGridCoaches.Rows + 1
                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 0) = cmbCoachType.ItemData(cmbCoachType.ListIndex)
                flexGridCoaches.TextMatrix(flexGridCoaches.Rows - 1, 1) = txtCoachNo
                MsgBox "Record Successfully Saved.", vbInformation
                cmdeditSave.Visible = False
                cmdDeleteRow.Visible = False
            Else
                MsgBox "Coach No. Already Added", vbCritical
            End If
        Else
            MsgBox "Length Of The Coach No. Must Have 10 Digits", vbCritical
        End If
    Else
        MsgBox "Please Fill all Fields.", vbCritical
    End If
End Sub

Private Sub flexGridCoaches_Click()
    If flexGridCoaches.Row > 0 Then
        txtCoachNo.Text = flexGridCoaches.TextMatrix(flexGridCoaches.Row, 1)
        cmdSave.Visible = False
        cmdeditSave.Visible = True
        cmdDeleteRow.Visible = True
        cmdDeleteRow.Top = flexGridCoaches.CellTop + flexGridCoaches.Top
    End If
End Sub

Private Sub Form_load()
    Label3.Caption = " Select Coach Type For Adding And Updating Record."
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
    flexGridCoaches.Cols = 2
    flexGridCoaches.FixedCols = 1
    flexGridCoaches.TextMatrix(0, 0) = "Coach ID"
    flexGridCoaches.TextMatrix(0, 1) = "Coach No."
    flexGridCoaches.ColWidth(0) = 1200
    flexGridCoaches.ColWidth(1) = 2500
    cmdeditSave.Visible = False
    cmdDeleteRow.Visible = False
End Sub

Private Sub Timer1_Timer()
    strs = Mid(Label3.Caption, 1, 1)
    Label3.Caption = Mid(Label3.Caption, 2, Len(Label3.Caption)) & strs
End Sub


Private Sub txtCoachNo_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(txtCoachNo.Text) < 10 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            txtCoachNo.Text = txtCoachNo.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        txtCoachNo.Text = txtCoachNo.Text & Chr(KeyAscii)
    End If
End If
End Sub
