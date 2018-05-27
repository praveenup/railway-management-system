VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRegionDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Region Search"
   ClientHeight    =   4680
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Bodoni MT"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRegion 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin MSDataListLib.DataList dataList 
      Height          =   3120
      Left            =   840
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmRegionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstRegion As ADODB.Recordset
Public regionID As Long
Option Explicit

Private Sub dataList_DblClick()
    If DataList.BoundText <> "" Then
        regionID = DataList.BoundText
        frmRegion.txtRegion.Text = ""
        frmRegion.txtRegion.Text = DataList.Text
        frmRegion.txtRegion.Enabled = True
        frmRegion.txtRegion.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmRegionDialog.Height = 1600
End Sub

Private Sub txtRegion_Change()
    Set rstRegion = New ADODB.Recordset
    rstRegion.CursorLocation = adUseClient
    rstRegion.Open "select * from region where region like '" & "%" & txtRegion.Text & "%' ", railCn
    If rstRegion.RecordCount > 0 Then
        Set DataList.RowSource = rstRegion
        Set DataList.DataSource = rstRegion
        DataList.BoundColumn = rstRegion.Fields(0).Name
        DataList.ListField = rstRegion.Fields(1).Name
        frmRegionDialog.Height = 5500
    End If
End Sub

Private Sub txtRegion_KeyPress(KeyAscii As Integer)
Call validation(2, KeyAscii, txtRegion)
End Sub
