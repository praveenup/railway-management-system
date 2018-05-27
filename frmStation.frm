VERSION 5.00
Begin VB.Form frmStation 
   BackColor       =   &H8000000E&
   Caption         =   "Station"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   12630
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   240
      Top             =   2040
   End
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   2880
      Picture         =   "frmStation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   8280
      Picture         =   "frmStation.frx":2742
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   4680
      Picture         =   "frmStation.frx":5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   6480
      Picture         =   "frmStation.frx":78BA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7080
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
      Height          =   5775
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   8415
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   6840
         Picture         =   "frmStation.frx":A1B8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cmbRegion 
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
         ItemData        =   "frmStation.frx":BABA
         Left            =   3600
         List            =   "frmStation.frx":BABC
         TabIndex        =   3
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox txtStn 
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
         Left            =   3600
         TabIndex        =   1
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtPlateform 
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
         Left            =   3600
         TabIndex        =   2
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox txtStnCode 
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
         Left            =   3600
         TabIndex        =   0
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Region:"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Number of Plateforms:"
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
         Left            =   1320
         TabIndex        =   7
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label2 
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
         Left            =   1320
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Station Code:"
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
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
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
      Left            =   960
      TabIndex        =   15
      Top             =   600
      Width           =   12615
   End
   Begin VB.Image Image4 
      Height          =   420
      Left            =   11400
      Picture         =   "frmStation.frx":BABE
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   5985
      Left            =   10560
      Picture         =   "frmStation.frx":BF98
      Top             =   1080
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   5985
      Left            =   1320
      Picture         =   "frmStation.frx":1C5B2
      Top             =   1080
      Width           =   825
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   360
      Picture         =   "frmStation.frx":2CBCC
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD STATION INFORMATION"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmStation.frx":2D517
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstStn As ADODB.Recordset
Dim rstStn1 As ADODB.Recordset
Dim rstRegion As ADODB.Recordset
Dim cmdstn As ADODB.Command
Dim ID As Long
Dim i As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
    If saveUpdate = 2 Then
            If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
                Set cmdstn = New ADODB.Command
                cmdstn.CommandType = adCmdText
                cmdstn.ActiveConnection = railCn
                cmdstn.CommandText = "delete from station where stnid=" & frmStationDialog.stnId & ""
                cmdstn.Execute
                MsgBox "Record Successfully Deleted", vbInformation
                txtStnCode.Enabled = False
                txtStn.Enabled = False
                cmbRegion.Enabled = False
                txtPlateform.Enabled = False
                txtStnCode.Text = ""
                txtStn.Text = ""
                cmbRegion.ListIndex = -1
                txtPlateform.Text = ""
                cmbRegion.Text = ""
                txtStnCode.BackColor = vbButtonFace
                txtStn.BackColor = vbButtonFace
                txtPlateform.BackColor = vbButtonFace
                cmbRegion.BackColor = vbButtonFace
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
    txtStnCode.Enabled = True
    txtStn.Enabled = True
    txtPlateform.Enabled = True
    cmbRegion.Enabled = True
    txtStnCode.BackColor = vbHighlightText
    txtStn.BackColor = vbHighlightText
    txtPlateform.BackColor = vbHighlightText
    cmbRegion.BackColor = vbHighlightText
    saveUpdate = 1
    txtStnCode.Text = ""
    txtStn.Text = ""
    txtPlateform.Text = ""
    cmbRegion.ListIndex = -1
End Sub
Private Function checkAlready() As Boolean
    Set rstStn1 = New ADODB.Recordset
    rstStn1.CursorLocation = adUseClient
    rstStn1.Open "select * from station ", railCn
    If rstStn1.RecordCount > 0 Then
        rstStn1.MoveFirst
        For i = 0 To rstStn1.RecordCount - 1
            If rstStn1("stncode") = txtStnCode.Text Or rstStn1("stnname") = txtStn.Text Then
                checkAlready = False
                rstStn1.Close
                Exit Function
            End If
            rstStn1.MoveNext
        Next
    End If
    checkAlready = True
    rstStn1.Close
End Function
Private Function checkAlready1() As Boolean
    Set rstStn1 = New ADODB.Recordset
    rstStn1.CursorLocation = adUseClient
    rstStn1.Open "select * from station ", railCn
    If rstStn1.RecordCount > 0 Then
        rstStn1.MoveFirst
        For i = 0 To rstStn1.RecordCount - 1
            If (rstStn1("stncode") = txtStnCode.Text Or rstStn1("stnname") = txtStn.Text) And (txtStn.Text <> frmStationDialog.stnName Or txtStnCode.Text <> frmStationDialog.stnCode) Then
                checkAlready1 = False
                rstStn1.Close
                Exit Function
            End If
            rstStn1.MoveNext
        Next
    End If
    checkAlready1 = True
    rstStn1.Close
End Function
Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
        If txtStnCode <> "" And txtStn <> "" And txtPlateform <> "" And cmbRegion.Text <> "" Then
                Set cmdstn = New ADODB.Command
                Set rstStn = New ADODB.Recordset
                cmdstn.CommandType = adCmdText
                cmdstn.ActiveConnection = railCn
                If saveUpdate = 1 Then
                    If checkAlready() Then
                        rstStn.Open "select max(stnID) from station", railCn
                        If rstStn.Fields(0) > 0 Then
                            ID = rstStn.Fields(0) + 1
                        Else
                            ID = 1
                        End If
                        cmdstn.CommandText = "insert into station values(" & ID & ",'" & txtStnCode & "','" & txtStn & "'," & txtPlateform & "," & cmbRegion.ItemData(cmbRegion.ListIndex) & ")"
                        cmdstn.Execute
                        saveUpdate = 0
                        MsgBox "Record Successfully Saved", vbInformation
                    Else
                        MsgBox "Station Information Already Exists, Please Give Another StnCode or StnName", vbCritical
                        Exit Sub
                    End If
                ElseIf saveUpdate = 2 Then
                    If checkAlready1() Then
                        cmdstn.CommandText = "update station set stnCode='" & txtStnCode.Text & "',stnName='" & txtStn & "',plateforms=" & txtPlateform & " where stnID = " & frmStationDialog.stnId & " "
                        cmdstn.Execute
                        If cmbRegion.ListIndex <> -1 Then
                            cmdstn.CommandText = "update  station set regionID=" & cmbRegion.ItemData(cmbRegion.ListIndex) & " where stnID = " & frmStationDialog.stnId & " "
                            cmdstn.Execute
                        End If
                        saveUpdate = 0
                        MsgBox "Record Successfully Updated", vbInformation
                    Else
                        MsgBox "Station Information Already Exists, Please Give Another StnCode or StnName", vbCritical
                        Exit Sub
                    End If
                End If
                txtStnCode.Enabled = False
                txtStn.Enabled = False
                cmbRegion.Enabled = False
                txtPlateform.Enabled = False
                txtStnCode.Text = ""
                txtStn.Text = ""
                cmbRegion.ListIndex = -1
                txtPlateform.Text = ""
                cmbRegion.Text = ""
                txtStnCode.BackColor = vbButtonFace
                txtStn.BackColor = vbButtonFace
                txtPlateform.BackColor = vbButtonFace
                cmbRegion.BackColor = vbButtonFace
        Else
            MsgBox "Please Fill all Fields", vbCritical
        End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub


Private Sub cmdSearch_Click()
    frmStationDialog.Show 1
End Sub

Private Sub Form_load()
    Set rstRegion = New ADODB.Recordset
    rstRegion.CursorLocation = adUseClient
    rstRegion.Open "select * from region", railCn
    If rstRegion.RecordCount > 1 Then
        i = 0
        rstRegion.MoveFirst
        Do While Not rstRegion.EOF
            cmbRegion.AddItem rstRegion(1)
            cmbRegion.ItemData(i) = rstRegion(0)
            rstRegion.MoveNext
            i = i + 1
        Loop
    End If
    rstRegion.Close
    txtStnCode.Enabled = False
    txtStn.Enabled = False
    txtPlateform.Enabled = False
    Label9.Caption = " Please click Add New Button to Add New Record (OR) Search and Select the Record for Updating Existing Record (OR)"
    cmbRegion.Enabled = False
End Sub


Private Sub Timer2_Timer()
    strs = Mid(Label9.Caption, 1, 1)
    Label9.Caption = Mid(Label9.Caption, 2, Len(Label9.Caption)) & strs
End Sub

Private Sub txtPlateform_KeyPress(KeyAscii As Integer)
    Call validation(1, KeyAscii, txtPlateform)
End Sub


Private Sub txtStn_KeyPress(KeyAscii As Integer)
      Call validation(2, KeyAscii, txtStn)
End Sub


Private Sub txtStnCode_KeyPress(KeyAscii As Integer)
      Call validation(3, KeyAscii, txtStnCode)
End Sub
