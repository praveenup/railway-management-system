VERSION 5.00
Begin VB.Form frmBerth 
   BackColor       =   &H8000000E&
   Caption         =   "Berth Type"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   8595
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
      Height          =   1935
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   7815
      Begin VB.TextBox txtBerth 
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
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   5880
         Picture         =   "frmBerth.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Berth Name:"
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
         Left            =   360
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   4440
      Picture         =   "frmBerth.frx":1902
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   2640
      Picture         =   "frmBerth.frx":4200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   6240
      Picture         =   "frmBerth.frx":6A7A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   840
      Picture         =   "frmBerth.frx":9378
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   11400
      Picture         =   "frmBerth.frx":BABA
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   2160
      Picture         =   "frmBerth.frx":BF94
      Top             =   480
      Width           =   5985
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   480
      Picture         =   "frmBerth.frx":1C1A6
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD BERTH TYPES"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmBerth.frx":1C98E
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmBerth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdBerth As ADODB.Command
Dim rstBerth As ADODB.Recordset
Dim ID As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
If saveUpdate = 2 Then
        If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
            Set cmdBerth = New ADODB.Command
            cmdBerth.CommandType = adCmdText
            cmdBerth.ActiveConnection = railCn
            cmdBerth.CommandText = "delete from Berthtype where BerthID=" & frmBerthDialog.berthID & ""
            cmdBerth.Execute
            MsgBox "Record Successfully Deleted", vbInformation
            txtBerth.Enabled = False
            txtBerth.Text = ""
            txtBerth.BackColor = vbButtonFace
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
    txtBerth.Enabled = True
    txtBerth.BackColor = vbHighlightText
    saveUpdate = 1
    txtBerth.Text = ""
End Sub

Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
    If txtBerth <> "" Then
        Set cmdBerth = New ADODB.Command
            Set rstBerth = New ADODB.Recordset
            cmdBerth.CommandType = adCmdText
            cmdBerth.ActiveConnection = railCn
            If saveUpdate = 1 Then
                rstBerth.Open "select max(BerthID) from berthtype", railCn
                If rstBerth.Fields(0) > 0 Then
                    ID = rstBerth.Fields(0) + 1
                Else
                    ID = 1
                End If
                    
                cmdBerth.CommandText = "insert into Berthtype values(" & ID & ",'" & txtBerth & "')"
                cmdBerth.Execute
                saveUpdate = 0
                MsgBox "Record Successfully Saved", vbInformation
            ElseIf saveUpdate = 2 Then
                cmdBerth.CommandText = "update Berthtype set Berthtype='" & txtBerth.Text & "' where BerthID=" & frmBerthDialog.berthID & ""
                cmdBerth.Execute
                saveUpdate = 0
                MsgBox "Record Successfully Updated", vbInformation
            End If
                txtBerth.Enabled = False
                txtBerth.Text = ""
                txtBerth.BackColor = vbButtonFace
    Else
        MsgBox "Please Fill all Fields", vbCritical
    End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub


Private Sub cmdSearch_Click()
frmBerthDialog.Show 1
End Sub

Private Sub Form_load()
    txtBerth.Enabled = False
    txtBerth.BackColor = vbButtonFace
End Sub



Private Sub txtBerth_KeyPress(KeyAscii As Integer)
      Call validation(2, KeyAscii, txtBerth)
End Sub
