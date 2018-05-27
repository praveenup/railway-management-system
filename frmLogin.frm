VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Height          =   375
      Left            =   4200
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdLogin 
      Height          =   375
      Left            =   1560
      Picture         =   "frmLogin.frx":319E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtpass 
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
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select User Type"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   2520
      Width           =   5415
      Begin VB.OptionButton optAdmin 
         BackColor       =   &H80000005&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optClerk 
         BackColor       =   &H8000000E&
         Caption         =   "Clerk"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.Image Image5 
         Height          =   720
         Left            =   2400
         Picture         =   "frmLogin.frx":633C
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.TextBox txtServiceNo 
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
      Left            =   3360
      TabIndex        =   1
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   1200
      Picture         =   "frmLogin.frx":6A7D
      Top             =   6240
      Width           =   5985
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   2880
      Picture         =   "frmLogin.frx":16C8F
      Top             =   480
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Image Image6 
      Height          =   720
      Left            =   720
      Picture         =   "frmLogin.frx":17683
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "USER LOGIN "
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
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   0
      Picture         =   "frmLogin.frx":17FB8
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdLogin_Click()
Set rs = New ADODB.Recordset
If txtServiceNo.Text <> "" And (optAdmin <> False Or optClerk <> False) Then
    If optClerk.Value = True Then
        opt = "Clerk"
    Else
        opt = "Admin"
    End If
    rs.CursorLocation = adUseClient
    rs.Open "select * from useraccount where serviceno='" & txtServiceNo & " ' and usertype='" & opt & "'", railCn
    If rs.RecordCount > 0 Then
        If txtpass.Text = rs.Fields("passw") Then
            userAccountID = rs("accountID")
            frmMDI.Show
            Unload Me
        Else
            MsgBox "Password is Incorrect", vbCritical
            txtServiceNo.SetFocus
        End If
    Else
        MsgBox "User Account Doesn't Exists", vbCritical
    End If
    rs.Close
 Else
    MsgBox "Please Fill All Fields", vbCritical
End If
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
