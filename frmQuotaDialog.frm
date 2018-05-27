VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmQuotaDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Quota"
   ClientHeight    =   4740
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtQuota 
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
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin MSDataListLib.DataList dataList 
      Height          =   3120
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmQuotaDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstQuota As ADODB.Recordset
Public quotaID As Long
Option Explicit

Private Sub dataList_DblClick()
    If DataList.BoundText <> "" Then
        quotaID = DataList.BoundText
        frmQuota.txtQuota.Text = ""
        frmQuota.txtQuota.Text = DataList.Text
        frmQuota.txtQuota.Enabled = True
        frmQuota.txtQuota.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmQuotaDialog.Height = 1365
End Sub

Private Sub txtQuota_Change()
    Set rstQuota = New ADODB.Recordset
    rstQuota.CursorLocation = adUseClient
    rstQuota.Open "select * from quota where quotaName like '" & "%" & txtQuota.Text & "%'", railCn
    If rstQuota.RecordCount > 0 Then
        Set DataList.RowSource = rstQuota
        Set DataList.DataSource = rstQuota
        DataList.BoundColumn = rstQuota.Fields(0).Name
        DataList.ListField = rstQuota.Fields(1).Name
        frmQuotaDialog.Height = 5140
    End If
End Sub

Private Sub txtQuota_KeyPress(KeyAscii As Integer)
      Call validation(2, KeyAscii, txtQuota)
End Sub
