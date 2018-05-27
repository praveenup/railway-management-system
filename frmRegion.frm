VERSION 5.00
Begin VB.Form frmRegion 
   BackColor       =   &H8000000E&
   Caption         =   "Region"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   9105
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   4680
      Picture         =   "frmRegion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   2880
      Picture         =   "frmRegion.frx":28FE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   6480
      Picture         =   "frmRegion.frx":5178
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   1080
      Picture         =   "frmRegion.frx":7A76
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
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
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   7815
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   5760
         Picture         =   "frmRegion.frx":A1B8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtRegion 
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
         Caption         =   "Region Name:"
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
      Picture         =   "frmRegion.frx":BABA
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   2400
      Picture         =   "frmRegion.frx":BF94
      Top             =   600
      Width           =   5985
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   480
      Picture         =   "frmRegion.frx":1C1A6
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD REGION"
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
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmRegion.frx":1C98E
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdRegion As ADODB.Command
Dim rstRegion As ADODB.Recordset
Dim ID As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
    If saveUpdate = 2 Then
        If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
            Set cmdRegion = New ADODB.Command
            cmdRegion.CommandType = adCmdText
            cmdRegion.ActiveConnection = railCn
            cmdRegion.CommandText = "delete from region where regionID=" & frmRegionDialog.regionID & ""
            cmdRegion.Execute
            MsgBox "Record Successfully Deleted", vbInformation
            txtRegion.Enabled = False
            txtRegion.Text = ""
            txtRegion.BackColor = vbButtonFace
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
    txtRegion.Enabled = True
    txtRegion.BackColor = vbHighlightText
    saveUpdate = 1
    txtRegion.Text = ""
End Sub

Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
    If txtRegion.Text <> "" Then
            Set cmdRegion = New ADODB.Command
            Set rstRegion = New ADODB.Recordset
            cmdRegion.CommandType = adCmdText
            cmdRegion.ActiveConnection = railCn
            If saveUpdate = 1 Then
                rstRegion.Open "select max(regionID) from region", railCn
                If rstRegion.Fields(0) > 0 Then
                    ID = rstRegion.Fields(0) + 1
                Else
                    ID = 1
                End If
                        
                cmdRegion.CommandText = "insert into region values(" & ID & ",'" & txtRegion & "')"
                cmdRegion.Execute
                saveUpdate = 0
                MsgBox "Record Successfully Saved", vbInformation
            ElseIf saveUpdate = 2 Then
                cmdRegion.CommandText = "update region set region='" & txtRegion.Text & "' where regionid=" & frmRegionDialog.regionID & " "
                cmdRegion.Execute
                saveUpdate = 0
                MsgBox "Record Successfully Updated", vbInformation
            End If
                txtRegion.Enabled = False
                txtRegion.Text = ""
                txtRegion.BackColor = vbButtonFace
        Else
            MsgBox "Please Fill all Fields", vbCritical
        End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub


Private Sub cmdSearch_Click()
frmRegionDialog.Show 1
End Sub

Private Sub Form_load()
    txtRegion.Enabled = False
    txtRegion.BackColor = vbButtonFace
End Sub




Private Sub txtRegion_KeyPress(KeyAscii As Integer)
     Call validation(2, KeyAscii, txtRegion)
End Sub
