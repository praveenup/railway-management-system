VERSION 5.00
Begin VB.Form frmQuota 
   BackColor       =   &H8000000E&
   Caption         =   "Quota"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   8535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   840
      Picture         =   "frmQuota.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   6240
      Picture         =   "frmQuota.frx":2742
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   2640
      Picture         =   "frmQuota.frx":5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   4440
      Picture         =   "frmQuota.frx":78BA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
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
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   7815
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   5880
         Picture         =   "frmQuota.frx":A1B8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtQuota 
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
         Left            =   2640
         TabIndex        =   0
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Quota Name:"
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
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   11400
      Picture         =   "frmQuota.frx":BABA
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   2160
      Picture         =   "frmQuota.frx":BF94
      Top             =   480
      Width           =   5985
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   360
      Picture         =   "frmQuota.frx":1C1A6
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD QUOTA"
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
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmQuota.frx":1CA02
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmQuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdQuota As ADODB.Command
Dim rstQuota As ADODB.Recordset
Dim ID As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
    If saveUpdate = 2 Then
        If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
            Set cmdQuota = New ADODB.Command
            cmdQuota.CommandType = adCmdText
            cmdQuota.ActiveConnection = railCn
            cmdQuota.CommandText = "delete from quota where quotaID=" & frmQuotaDialog.quotaID & ""
            cmdQuota.Execute
            MsgBox "Record Successfully Deleted", vbInformation
            txtQuota.Enabled = False
            txtQuota.Text = ""
            txtQuota.BackColor = vbButtonFace
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
    txtQuota.Enabled = True
    txtQuota.BackColor = vbHighlightText
    saveUpdate = 1
    txtQuota.Text = ""
End Sub

Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
    If txtQuota <> "" Then
        Set cmdQuota = New ADODB.Command
            Set rstQuota = New ADODB.Recordset
            cmdQuota.CommandType = adCmdText
            cmdQuota.ActiveConnection = railCn
            If saveUpdate = 1 Then
                rstQuota.Open "select max(quotaID) from quota", railCn
                If rstQuota.Fields(0) > 0 Then
                    ID = rstQuota.Fields(0) + 1
                Else
                    ID = 1
                End If
                    
                cmdQuota.CommandText = "insert into quota values(" & ID & ",'" & txtQuota & "')"
                cmdQuota.Execute
                saveUpdate = 0
                MsgBox "Record Successfully Saved", vbInformation
            ElseIf saveUpdate = 2 Then
                cmdQuota.CommandText = "update quota set quotaName='" & txtQuota.Text & "' where quotaid=" & frmQuotaDialog.quotaID & ""
                cmdQuota.Execute
                saveUpdate = 0
                MsgBox "Record Successfully Updated", vbInformation
            End If
                txtQuota.Enabled = False
                txtQuota.Text = ""
                txtQuota.BackColor = vbButtonFace
    Else
        MsgBox "Please Fill all Fields", vbCritical
    End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub


Private Sub cmdSearch_Click()
frmQuotaDialog.Show 1
End Sub

Private Sub Form_load()
    txtQuota.Enabled = False
    txtQuota.BackColor = vbButtonFace
End Sub

Private Sub txtQuota_KeyPress(KeyAscii As Integer)
      Call validation(2, KeyAscii, txtQuota)
End Sub
