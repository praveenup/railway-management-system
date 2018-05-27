VERSION 5.00
Begin VB.Form frmTrainType 
   BackColor       =   &H8000000E&
   Caption         =   "Train Type"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14145
   BeginProperty Font 
      Name            =   "Bodoni MT"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   14145
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   360
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   -2520
      Top             =   4920
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Picture         =   "frmTrainType.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Picture         =   "frmTrainType.frx":28FE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      Picture         =   "frmTrainType.frx":5178
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdNew 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Picture         =   "frmTrainType.frx":7A76
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
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
      Height          =   3855
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   8535
      Begin VB.OptionButton opNo 
         BackColor       =   &H8000000E&
         Caption         =   "No"
         Height          =   255
         Left            =   5520
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton opYes 
         BackColor       =   &H80000005&
         Caption         =   "Yes"
         Height          =   255
         Left            =   3600
         TabIndex        =   1
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         Picture         =   "frmTrainType.frx":A1B8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtType 
         BackColor       =   &H8000000F&
         Height          =   420
         Left            =   3600
         TabIndex        =   0
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtSpeed 
         BackColor       =   &H8000000F&
         Height          =   420
         Left            =   3600
         TabIndex        =   3
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Catering Availability:"
         Height          =   495
         Left            =   960
         TabIndex        =   8
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Type Name:"
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Maximum Speed:"
         Height          =   495
         Left            =   960
         TabIndex        =   5
         Top             =   3000
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
      Left            =   360
      TabIndex        =   14
      Top             =   1440
      Width           =   12615
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   11400
      Picture         =   "frmTrainType.frx":BABA
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   3120
      Picture         =   "frmTrainType.frx":BF94
      Top             =   480
      Width           =   5985
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   600
      Picture         =   "frmTrainType.frx":1C1A6
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD TRAIN TYPES"
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
      TabIndex        =   13
      Top             =   0
      Width           =   3735
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmTrainType.frx":1CAF1
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmTrainType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdTrainType As ADODB.Command
Dim rstTrainType As ADODB.Recordset
Dim ID As Long  'for typeID
Dim opt As Boolean 'option button selection

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
If saveUpdate = 2 Then
        If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
            Set cmdTrainType = New ADODB.Command
            cmdTrainType.CommandType = adCmdText
            cmdTrainType.ActiveConnection = railCn
            cmdTrainType.CommandText = "delete from traintype where typeid=" & frmTrainTypeDialog.typeID & ""
            cmdTrainType.Execute
            MsgBox "Record Successfully Deleted", vbInformation
            txtType.Enabled = False
            txtSpeed.Enabled = False
            opYes.Enabled = False
            opNo.Enabled = False
            txtType.Text = ""
            txtSpeed.Text = ""
            opYes = False
            opNo = False
            txtType.BackColor = vbButtonFace
            txtSpeed.BackColor = vbButtonFace
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
    txtType.Enabled = True
    txtSpeed.Enabled = True
    opYes.Enabled = True
    opNo.Enabled = True
    txtType.BackColor = vbHighlightText
    txtSpeed.BackColor = vbHighlightText
    saveUpdate = 1
    txtType.Text = ""
    txtSpeed.Text = ""
End Sub

Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
    If txtType.Text <> "" And txtSpeed.Text <> "" And (opYes <> False Or opNo <> False) Then
            Set cmdTrainType = New ADODB.Command
            Set rstTrainType = New ADODB.Recordset
            cmdTrainType.CommandType = adCmdText
            cmdTrainType.ActiveConnection = railCn
            If saveUpdate = 1 Then
                rstTrainType.Open "select max(typeID) from traintype", railCn
                If rstTrainType.Fields(0) > 0 Then
                    ID = rstTrainType.Fields(0) + 1
                Else
                    ID = 1
                End If
                If opYes = True Then
                    opt = True
                Else
                    opt = False
                End If
                cmdTrainType.CommandText = "insert into traintype values(" & ID & ",'" & txtType & "'," & opt & ",'" & txtSpeed & "')"
                cmdTrainType.Execute
                saveUpdate = 0
                MsgBox "Record Successfully Saved", vbInformation
            ElseIf saveUpdate = 2 Then
                If opYes = True Then
                    opt = True
                Else
                    opt = False
                End If
                cmdTrainType.CommandText = "update trainType set typeName='" & txtType.Text & "',catering=" & opt & ",maxSpeed=" & txtSpeed & " where typeID = " & frmTrainTypeDialog.typeID & " "
                cmdTrainType.Execute
                saveUpdate = 0
                MsgBox "Record Successfully Updated", vbInformation
            End If
                txtType.Enabled = False
                txtSpeed.Enabled = False
                opYes.Enabled = False
                opNo.Enabled = False
                txtType.Text = ""
                txtSpeed.Text = ""
                opYes = False
                opNo = False
                txtType.BackColor = vbButtonFace
                txtSpeed.BackColor = vbButtonFace
    Else
        MsgBox "Please Fill all Fields", vbCritical
    End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub


Private Sub cmdSearch_Click()
frmTrainTypeDialog.Show 1
End Sub

Private Sub Form_load()
    txtType.Enabled = False
    txtSpeed.Enabled = False
    opYes.Enabled = False
    opNo.Enabled = False
    Label9.Caption = " Please click Add New Button to Add New Record (OR) Search and Select the Record for Updating Existing Record."

End Sub




Private Sub Timer2_Timer()
    strs = Mid(Label9.Caption, 1, 1)
    Label9.Caption = Mid(Label9.Caption, 2, Len(Label9.Caption)) & strs
End Sub

Private Sub txtSpeed_KeyPress(KeyAscii As Integer)
    Call validation(1, KeyAscii, txtSpeed)
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
     Call validation(2, KeyAscii, txtType)
End Sub
