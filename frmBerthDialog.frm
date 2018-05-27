VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBerthDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Berth Type"
   ClientHeight    =   4950
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBerth 
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
      TabIndex        =   1
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
      Caption         =   "Berth Type:"
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
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmBerthDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstBerth As ADODB.Recordset
Public berthID As Long
Option Explicit

Private Sub dataList_DblClick()
    If DataList.BoundText <> "" Then
        berthID = DataList.BoundText
        frmBerth.txtBerth.Text = ""
        frmBerth.txtBerth.Text = DataList.Text
        frmBerth.txtBerth.Enabled = True
        frmBerth.txtBerth.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmBerthDialog.Height = 1365
End Sub

Private Sub txtBerth_Change()
    Set rstBerth = New ADODB.Recordset
    rstBerth.CursorLocation = adUseClient
    rstBerth.Open "select * from berthtype where berthtype like '" & "%" & txtBerth.Text & "%'", railCn
    If rstBerth.RecordCount > 0 Then
        Set DataList.RowSource = rstBerth
        Set DataList.DataSource = rstBerth
        DataList.BoundColumn = rstBerth.Fields(0).Name
        DataList.ListField = rstBerth.Fields(1).Name
        frmBerthDialog.Height = 5140
    End If
End Sub


Private Sub txtBerth_KeyPress(KeyAscii As Integer)
Call validation(2, KeyAscii, txtBerth)
End Sub
