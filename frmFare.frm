VERSION 5.00
Begin VB.Form frmFare 
   BackColor       =   &H8000000E&
   Caption         =   "Fare"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   12315
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   1800
      Picture         =   "frmFare.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   7200
      Picture         =   "frmFare.frx":2742
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   3600
      Picture         =   "frmFare.frx":5040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   5400
      Picture         =   "frmFare.frx":78BA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
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
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   8895
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   6720
         Picture         =   "frmFare.frx":A1B8
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
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
         Left            =   3480
         TabIndex        =   6
         Top             =   1920
         Width           =   2895
      End
      Begin VB.ComboBox cmbCoachType 
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
         TabIndex        =   5
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox txtFare 
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
         TabIndex        =   1
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label Label1 
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
         Left            =   1080
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
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
         TabIndex        =   3
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Fare (per KM):"
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
         TabIndex        =   2
         Top             =   3960
         Width           =   2055
      End
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
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   12615
   End
   Begin VB.Image Image4 
      Height          =   5985
      Left            =   10080
      Picture         =   "frmFare.frx":BABA
      Top             =   1200
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   5985
      Left            =   360
      Picture         =   "frmFare.frx":1C0D4
      Top             =   1200
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11400
      Picture         =   "frmFare.frx":2C6EE
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   720
      Picture         =   "frmFare.frx":2CBC8
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD FARE INFORMATION"
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
      Left            =   1440
      TabIndex        =   12
      Top             =   0
      Width           =   4215
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmFare.frx":2D513
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmFare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstFare As ADODB.Recordset
Dim rstFare1 As ADODB.Recordset
Dim rstTrainType As ADODB.Recordset
Dim rstCoachType As ADODB.Recordset
Dim cmdFare As ADODB.Command
Dim ID As Long
Dim i As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
    If saveUpdate = 2 Then
            If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
                Set cmdFare = New ADODB.Command
                cmdFare.CommandType = adCmdText
                cmdFare.ActiveConnection = railCn
                cmdFare.CommandText = "delete from station where typeid=" & frmfareDialog.typeId & " and coachid=" & frmfareDialog.coachID & ""
                cmdFare.Execute
                MsgBox "Record Successfully Deleted", vbInformation
                txtFare.Enabled = False
                cmbCoachType.Enabled = False
                cmbTrainType.Enabled = False
                txtFare.Text = ""
                cmbCoachType.ListIndex = -1
                cmbTrainType.ListIndex = -1
                cmbTrainType.BackColor = vbButtonFace
                cmbCoachType.BackColor = vbButtonFace
                txtFare.BackColor = vbButtonFace
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
    txtFare.Enabled = True
    cmbCoachType.Enabled = True
    cmbTrainType.Enabled = True
    txtFare.BackColor = vbHighlightText
    cmbTrainType.BackColor = vbHighlightText
    cmbCoachType.BackColor = vbHighlightText
    saveUpdate = 1
    txtFare.Text = ""
    cmbCoachType.ListIndex = -1
    cmbTrainType.ListIndex = -1
End Sub
Private Function checkAlready() As Boolean
    Set rstFare1 = New ADODB.Recordset
    rstFare1.CursorLocation = adUseClient
    rstFare1.Open "select * from fare ", railCn
    If rstFare1.RecordCount > 0 Then
        rstFare1.MoveFirst
        For i = 0 To rstFare1.RecordCount - 1
            If rstFare1("typeid") = cmbTrainType.ItemData(cmbTrainType.ListIndex) And rstFare1("coachid") = cmbCoachType.ItemData(cmbCoachType.ListIndex) Then
                checkAlready = False
                rstFare1.Close
                Exit Function
            End If
            rstFare1.MoveNext
        Next
    End If
    checkAlready = True
    rstFare1.Close
End Function
Private Function checkAlready1() As Boolean
    Set rstFare1 = New ADODB.Recordset
    rstFare1.CursorLocation = adUseClient
    rstFare1.Open "select * from fare ", railCn
    If rstFare1.RecordCount > 0 Then
        rstFare1.MoveFirst
        For i = 0 To rstFare1.RecordCount - 1
            If (rstFare1("typeid") = cmbTrainType.ItemData(cmbTrainType.ListIndex) And rstFare1("coachid") = cmbCoachType.ItemData(cmbCoachType.ListIndex)) And (rstFare1("typeid") <> frmfareDialog.typeId Or rstFare1("coachid") <> frmfareDialog.coachID) Then
                checkAlready1 = False
                rstFare1.Close
                Exit Function
            End If
            rstFare1.MoveNext
        Next
    End If
    checkAlready1 = True
    rstFare1.Close
End Function
Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
        If txtFare <> "" And cmbCoachType.ListIndex <> -1 And cmbTrainType.ListIndex <> -1 Then
                Set cmdFare = New ADODB.Command
                Set rstFare = New ADODB.Recordset
                cmdFare.CommandType = adCmdText
                cmdFare.ActiveConnection = railCn
                If saveUpdate = 1 Then
                    If checkAlready() Then
                        cmdFare.CommandText = "insert into fare values(" & cmbTrainType.ItemData(cmbTrainType.ListIndex) & "," & cmbCoachType.ItemData(cmbCoachType.ListIndex) & "," & txtFare & ")"
                        cmdFare.Execute
                        saveUpdate = 0
                        MsgBox "Record Successfully Saved", vbInformation
                    Else
                        MsgBox "Fare Information Already Exists.", vbCritical
                        Exit Sub
                    End If
                ElseIf saveUpdate = 2 Then
                    If checkAlready1() Then
                        cmdFare.CommandText = "update fare set typeid=" & cmbTrainType.ItemData(cmbTrainType.ListIndex) & ",coachid=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & ",fare=" & txtFare & " where coachid = " & frmfareDialog.coachID & " and typeid = " & frmfareDialog.typeId & " "
                        cmdFare.Execute
                        saveUpdate = 0
                        MsgBox "Record Successfully Updated", vbInformation
                    Else
                        MsgBox "Fare Information Already Exists.", vbCritical
                        Exit Sub
                    End If
                End If
                txtFare.Enabled = False
                cmbCoachType.Enabled = False
                cmbTrainType.Enabled = False
                txtFare.Text = ""
                cmbCoachType.ListIndex = -1
                cmbTrainType.ListIndex = -1
                cmbTrainType.BackColor = vbButtonFace
                cmbCoachType.BackColor = vbButtonFace
                txtFare.BackColor = vbButtonFace
        Else
            MsgBox "Please Fill all Fields", vbCritical
        End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub


Private Sub cmdSearch_Click()
    frmfareDialog.Show 1
End Sub



Private Sub Timer2_Timer()
    strs = Mid(Label9.Caption, 1, 1)
    Label9.Caption = Mid(Label9.Caption, 2, Len(Label9.Caption)) & strs
End Sub

Private Sub Form_load()
    Set rstTrainType = New ADODB.Recordset
    rstTrainType.CursorLocation = adUseClient
    rstTrainType.Open "select * from traintype", railCn
    If rstTrainType.RecordCount > 1 Then
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

    Set rstCoachType = New ADODB.Recordset
    rstCoachType.CursorLocation = adUseClient
    rstCoachType.Open "select * from coach", railCn
    If rstCoachType.RecordCount > 1 Then
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
    txtFare.Enabled = False
    cmbCoachType.Enabled = False
    cmbTrainType.Enabled = False
    Label9.Caption = " Please click Add New Button to Add New Record (OR) Search and Select the Record for Updating Existing Record (OR)"

End Sub







