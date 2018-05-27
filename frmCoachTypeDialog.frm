VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCoachTypeDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Coach"
   ClientHeight    =   5160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCoachType 
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
      Height          =   2910
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5133
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Coach Type:"
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
Attribute VB_Name = "frmCoachTypeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCoach As ADODB.Recordset
Dim rstCoachBerth As ADODB.Recordset
Public coachID As Long
Public coachTypeName As Variant
Public coachTypeInitial As Variant
Dim i As Integer
Option Explicit

Private Sub dataList_DblClick()
    If DataList.BoundText <> "" Then
        frmCoachType.Frame1.Height = 3000
        frmCoachType.Frame1.Width = 7215
        frmCoachType.txtCoachType.Text = ""
        frmCoachType.chkChair = 0
        frmCoachType.chkAC = 0
        frmCoachType.chkPassenger = 0
        frmCoachType.cmbCoachInitial.ListIndex = -1
        frmCoachType.txtCoachType.Enabled = True
        frmCoachType.cmbCoachInitial.Enabled = True
        frmCoachType.chkPassenger.Enabled = True
        frmCoachType.txtCoachType.BackColor = vbHighlightText
        frmCoachType.cmbCoachInitial.BackColor = vbHighlightText
        frmCoachType.cmdNew.Left = 960
        frmCoachType.cmdNew.Top = 4560
        frmCoachType.cmdSave.Left = 2760
        frmCoachType.cmdSave.Top = 4560
        frmCoachType.cmdDelete.Left = 4560
        frmCoachType.cmdDelete.Top = 4560
        frmCoachType.cmdClose.Left = 6360
        frmCoachType.cmdClose.Top = 4560
        For i = 0 To 11
            frmCoachType.cmbBox(i).Visible = False
            frmCoachType.lblBerth(i).Visible = False
            frmCoachType.cmbBox(i).ListIndex = -1
        Next
        coachID = DataList.BoundText
        frmCoachType.txtCoachType.Text = DataList.Text
        coachTypeName = DataList.Text
        Set rstCoach = New ADODB.Recordset
        rstCoach.CursorLocation = adUseClient
        rstCoach.Open "select * from coach where coachtypeid=" & coachID & "", railCn
        frmCoachType.cmbCoachInitial.Text = rstCoach("coachinitial")
        coachTypeInitial = rstCoach("coachinitial")
        If rstCoach("passengercoach") = True Then
            frmCoachType.cmdNew.Top = 7920
            frmCoachType.cmdSave.Top = 7920
            frmCoachType.cmdDelete.Top = 7920
            frmCoachType.cmdClose.Top = 7920
            frmCoachType.Frame1.Height = 6975
            frmCoachType.Frame1.Width = 16095
            If rstCoach("passengercoach") = True Then
                frmCoachType.chkPassenger = 1
            End If
            If rstCoach("ac") = True Then
                frmCoachType.chkAC = 1
            End If
            If rstCoach("chaircar") = True Then
                frmCoachType.chkChair = 1
            End If
            frmCoachType.txtCabins = rstCoach("cabin")
            frmCoachType.cmbCabinSeat.ListIndex = comboSearch(frmCoachType.cmbCabinSeat, rstCoach("cabinseats"))
            Set rstCoachBerth = New ADODB.Recordset
            rstCoachBerth.CursorLocation = adUseClient
            rstCoachBerth.Open "select * from coachberth where coachtypeid=" & coachID & "", railCn
                rstCoachBerth.MoveFirst
                For i = 0 To frmCoachType.cmbCabinSeat.ItemData(frmCoachType.cmbCabinSeat.ListIndex) - 1
                    frmCoachType.cmbBox(i).Visible = True
                    frmCoachType.lblBerth(i).Visible = True
                    frmCoachType.cmbBox(i).ListIndex = comboSearch(frmCoachType.cmbBox(i), rstCoachBerth("berthtypeid"))
                    rstCoachBerth.MoveNext
                Next
        End If
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmCoachTypeDialog.Height = 1365
End Sub

Private Sub txtCoachType_Change()
    Set rstCoach = New ADODB.Recordset
    rstCoach.CursorLocation = adUseClient
    rstCoach.Open "select * from coach where coachtypename like '" & "%" & txtCoachType.Text & "%'", railCn
    If rstCoach.RecordCount > 0 Then
        Set DataList.RowSource = rstCoach
        Set DataList.DataSource = rstCoach
        DataList.BoundColumn = rstCoach.Fields(0).Name
        DataList.ListField = rstCoach.Fields(1).Name
        frmCoachTypeDialog.Height = 5140
    End If
End Sub


Private Sub txtCoachType_KeyPress(KeyAscii As Integer)
Call validation(2, KeyAscii, txtCoachType)
End Sub

