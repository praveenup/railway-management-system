VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAccountDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Account"
   ClientHeight    =   5340
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
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
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DGAccount 
      CausesValidation=   0   'False
      Height          =   3735
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6588
      _Version        =   393216
      BackColor       =   -2147483634
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
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
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmAccountDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstAccount As ADODB.Recordset
Public accountID As Long
Private Sub dgaccount_dblClick()
    If DGAccount.Row <> -1 Then
        i = DGAccount.Row
        DGAccount.RowBookmark (i)
        frmAccount.txtServiceNo.Text = DGAccount.Columns(1)
        frmAccount.txtUser.Text = DGAccount.Columns(3)
        accountID = DGAccount.Columns(0)
        Set rstAccount = New ADODB.Recordset
        rstAccount.CursorLocation = adUseClient
        rstAccount.Open "select * from useraccount where accountId=" & accountID & "", railCn
        frmAccount.txtpass = rstAccount("passw")
        frmAccount.txtVerify = rstAccount("passw")
        rstAccount.Close
        If DGAccount.Columns(2) = "Admin" Then
            frmAccount.opAdmin = True
        Else
            frmAccount.opClerk = True
        End If
        
        frmAccount.txtServiceNo.Enabled = True
        frmAccount.txtUser.Enabled = True
        frmAccount.txtpass.Enabled = True
        frmAccount.txtVerify.Enabled = True
        frmAccount.opAdmin.Enabled = True
        frmAccount.opClerk.Enabled = True
        frmAccount.txtServiceNo.BackColor = vbHighlightText
        frmAccount.txtUser.BackColor = vbHighlightText
        frmAccount.txtpass.BackColor = vbHighlightText
        frmAccount.txtVerify.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmAccountDialog.Height = 1515
End Sub

Private Sub txtServiceNo_Change()
    Set rstAccount = New ADODB.Recordset
    rstAccount.CursorLocation = adUseClient
    rstAccount.Open "select accountID as AccountID,ServiceNo as ServiceNo,userType as Type,userName as  UserName from useraccount where serviceno like '%" & txtServiceNo & "%' ", railCn
    If rstAccount.RecordCount > 0 Then
        Set DGAccount.DataSource = rstAccount
        frmAccountDialog.Height = 5500
    Else
        frmAccountDialog.Height = 1515
    End If
End Sub




Private Sub txtServiceNo_KeyPress(KeyAscii As Integer)
Call validation(1, KeyAscii, txtServiceNo)
End Sub

