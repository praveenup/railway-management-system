VERSION 5.00
Begin VB.Form frmAccount 
   BackColor       =   &H8000000E&
   Caption         =   "User Account"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   14355
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNew 
      Height          =   495
      Left            =   1200
      Picture         =   "frmAccount.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Height          =   495
      Left            =   6600
      Picture         =   "frmAccount.frx":2742
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   3000
      Picture         =   "frmAccount.frx":5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   495
      Left            =   4800
      Picture         =   "frmAccount.frx":78BA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User Input Information"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   6135
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   8175
      Begin VB.OptionButton opAdmin 
         BackColor       =   &H80000005&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   5280
         Width           =   1695
      End
      Begin VB.OptionButton opClerk 
         BackColor       =   &H8000000E&
         Caption         =   "Clerk"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox txtServiceNo 
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
         Height          =   465
         Left            =   3240
         TabIndex        =   14
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   375
         Left            =   6480
         Picture         =   "frmAccount.frx":A1B8
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtUser 
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
         Height          =   465
         Left            =   3240
         TabIndex        =   4
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtPass 
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
         ForeColor       =   &H00000000&
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox txtVerify 
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
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Service No.:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "User Type:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Password:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Bodoni MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   2040
         Width           =   2295
      End
   End
   Begin VB.Image Image3 
      Height          =   5985
      Left            =   9240
      Picture         =   "frmAccount.frx":BABA
      Top             =   1080
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   5985
      Left            =   240
      Picture         =   "frmAccount.frx":1C0D4
      Top             =   1080
      Width           =   825
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   1440
      Picture         =   "frmAccount.frx":2C6EE
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "CREATE USER ACCOUNT"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11400
      Picture         =   "frmAccount.frx":2CFB0
      Top             =   0
      Width           =   11535
   End
   Begin VB.Image Image4 
      Height          =   420
      Left            =   0
      Picture         =   "frmAccount.frx":2D48A
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cmdAccount As ADODB.Command
Dim rstAccount As ADODB.Recordset
Dim ID As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo label
    If saveUpdate = 2 Then
        If vbYes = MsgBox("Are you sure want to Delete selected Record?", vbQuestion + vbYesNo, "") Then
            Set cmdAccount = New ADODB.Command
            cmdAccount.CommandType = adCmdText
            cmdAccount.ActiveConnection = railCn
            cmdAccount.CommandText = "delete from useraccount where accountID=" & frmAccountDialog.accountID & ""
            cmdAccount.Execute
            MsgBox "Record Successfully Deleted", vbInformation
            txtServiceNo.Enabled = False
            txtUser.Enabled = False
            txtpass.Enabled = False
            txtVerify.Enabled = False
            opClerk.Enabled = False
            opAdmin.Enabled = False
            txtServiceNo.Text = ""
            txtUser.Text = ""
            txtpass.Text = ""
            txtVerify.Text = ""
            opClerk.Value = False
            opAdmin.Value = False
            txtServiceNo.BackColor = vbButtonFace
            txtUser.BackColor = vbButtonFace
            txtpass.BackColor = vbButtonFace
            txtVerify.BackColor = vbButtonFace
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
    txtServiceNo.Enabled = True
    txtUser.Enabled = True
    txtpass.Enabled = True
    txtVerify.Enabled = True
    opClerk.Enabled = True
    opAdmin.Enabled = True
    txtServiceNo.BackColor = vbHighlightText
    txtUser.BackColor = vbHighlightText
    txtpass.BackColor = vbHighlightText
    txtVerify.BackColor = vbHighlightText
    opClerk.BackColor = vbHighlightText
    opAdmin.BackColor = vbHighlightText
    saveUpdate = 1
    txtServiceNo.Text = ""
    txtUser.Text = ""
    txtpass.Text = ""
    txtVerify.Text = ""
    opClerk.Value = False
    opAdmin.Value = False
End Sub

Private Function checkUserExists() As Boolean
    Set rstAccount = New ADODB.Recordset
    rstAccount.CursorLocation = adUseClient
    rstAccount.Open "select * from userAccount", railCn
    If rstAccount.RecordCount > 0 Then
        i = 0
        rstAccount.MoveFirst
        Do While Not rstAccount.EOF
            If txtUser = rstAccount(1) Then
            checkUserExists = False
            Exit Function
            End If
            rstAccount.MoveNext
            i = i + 1
        Loop
    End If
    rstAccount.Close
    checkUserExists = True
End Function

Private Sub cmdSave_Click()
If saveUpdate = 1 Or saveUpdate = 2 Then
    If Len(txtServiceNo.Text) = 10 Then
        If txtpass.Text = txtVerify.Text Then
            If txtServiceNo.Text <> "" And txtUser.Text <> "" And txtpass.Text <> "" And txtVerify.Text <> "" And (opAdmin <> False Or opClerk <> False) Then
                If checkUserExists() Then
                    Set cmdAccount = New ADODB.Command
                    Set rstAccount = New ADODB.Recordset
                    cmdAccount.CommandType = adCmdText
                    cmdAccount.ActiveConnection = railCn
                    If saveUpdate = 1 Then
                        rstAccount.Open "select max(accountID) from userAccount", railCn
                        If rstAccount.Fields(0) > 0 Then
                            ID = rstAccount.Fields(0) + 1
                        Else
                            ID = 1
                        End If
                        If opClerk.Value = True Then
                            opt = "Clerk"
                        Else
                            opt = "Admin"
                        End If
                        cmdAccount.CommandText = "insert into userAccount values(" & ID & ",'" & txtServiceNo & "','" & txtUser & "','" & txtpass & "','" & opt & "')"
                        cmdAccount.Execute
                        saveUpdate = 0
                        MsgBox "Record Successfully Saved", vbInformation
                    ElseIf saveUpdate = 2 Then
                        If opClerk.Value = True Then
                            opt = "Clerk"
                        Else
                            opt = "Admin"
                        End If
                        cmdAccount.CommandText = "update userAccount set   userName = '" & txtUser.Text & "',passw='" & txtpass & "' , userType = '" & opt & "' where accountID = " & frmAccountDialog.accountID & " "
                        Debug.Print "update userAccount set   userName = '" & txtUser.Text & "',passw='" & txtpass & "' , userType = '" & opt & "' where accountID = " & frmAccountDialog.accountID & " "
                        cmdAccount.Execute
                        saveUpdate = 0
                        MsgBox "Record Successfully Updated", vbInformation
                    End If
                    txtServiceNo.Enabled = False
                    txtUser.Enabled = False
                    txtpass.Enabled = False
                    txtVerify.Enabled = False
                    opClerk.Enabled = False
                    opAdmin.Enabled = False
                    txtServiceNo.Text = ""
                    txtUser.Text = ""
                    txtpass.Text = ""
                    txtVerify.Text = ""
                    opClerk.Value = False
                    opAdmin.Value = False
                    txtServiceNo.BackColor = vbButtonFace
                    txtUser.BackColor = vbButtonFace
                    txtpass.BackColor = vbButtonFace
                    txtVerify.BackColor = vbButtonFace
                Else
                    MsgBox "User Already Exists", vbCritical
                End If
            Else
                MsgBox "Please Fill all Fields", vbCritical
            End If
        Else
            MsgBox "Password Do Not Match, Please Re-Enter your password.", vbCritical
            txtpass.Text = ""
            txtVerify.Text = ""
        End If
    Else
        MsgBox "Service No. Must Have 10 Digits", vbCritical
    End If
Else
    MsgBox "Please click Add New Button to Add New Record OR Search and Select the Record for Updating Existing Record", vbCritical
End If
End Sub

Private Sub cmdSearch_Click()
frmAccountDialog.Show 1
End Sub

Private Sub Form_load()
    txtServiceNo.Enabled = False
    txtUser.Enabled = False
    txtpass.Enabled = False
    txtVerify.Enabled = False
    opClerk.Enabled = False
    opAdmin.Enabled = False
End Sub

Private Sub txtQuota_KeyPress(KeyAscii As Integer)
      Call validation(1, KeyAscii, txtQuota)
End Sub

Private Sub txtPass_Change()
    txtVerify.Text = ""
End Sub

Private Sub txtServiceNo_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 8 Then
    If Len(txtServiceNo.Text) < 10 Then
        If IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            txtServiceNo.Text = txtServiceNo.Text & Chr(KeyAscii)
        End If
    Else
        KeyAscii = 0
        txtServiceNo.Text = txtServiceNo.Text & Chr(KeyAscii)
    End If
End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    Call validation(2, KeyAscii, txtUser)
End Sub
