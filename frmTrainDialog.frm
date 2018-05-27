VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTrainDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Train"
   ClientHeight    =   5745
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTrain 
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
      Height          =   3885
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6853
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
      Caption         =   "Train Name:"
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
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmTrainDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTrain As ADODB.Recordset
Dim rstRoute As ADODB.Recordset
Dim rstTrainRoute As ADODB.Recordset
Dim rstdays As ADODB.Recordset
Public upTrainNo As Long
Const strChecked = "þ"
Const strUnChecked = "q"
Public downTrainNo As Long
Public trainName As Variant


Private Sub dataList_DblClick()
    If DataList.BoundText <> "" Then
        frmTrain.chkSun.Value = 0
        frmTrain.chkMon.Value = 0
        frmTrain.chkTue.Value = 0
        frmTrain.chkWed.Value = 0
        frmTrain.chkThr.Value = 0
        frmTrain.chkFri.Value = 0
        frmTrain.chkSat.Value = 0
        frmTrain.chkSun1.Value = 0
        frmTrain.chkMon1.Value = 0
        frmTrain.chkTue1.Value = 0
        frmTrain.chkWed1.Value = 0
        frmTrain.chkThr1.Value = 0
        frmTrain.chkFri1.Value = 0
        frmTrain.chkSat1.Value = 0
        upTrainNo = DataList.BoundText
        Set rstTrain = New ADODB.Recordset
        rstTrain.CursorLocation = adUseClient
        rstTrain.Open "select routeid,traintypeid,uptrainno,downtrainno from train where uptrainno= '" & upTrainNo & "' ", railCn
        If rstTrain.RecordCount > 0 Then
            downTrainNo = Val(rstTrain(3))
            frmTrain.txtUTrainNo.Text = rstTrain(2)
            frmTrain.txtDTrainNo.Text = rstTrain(3)
            trainName = DataList.Text
            frmTrain.txtTrainName.Text = DataList.Text
            frmTrain.cmbRoute.ListIndex = comboSearch(frmTrain.cmbRoute, rstTrain(0))
            frmTrain.cmbTrainType.ListIndex = comboSearch(frmTrain.cmbTrainType, rstTrain(1))
        End If
        
        Set rstdays = New ADODB.Recordset
        rstdays.CursorLocation = adUseClient
        rstdays.Open "select * from days where trainno= '" & upTrainNo & "' ", railCn
        If rstdays.RecordCount > 0 Then
            If rstdays(1) = True Then
                frmTrain.chkSun.Value = 1
            End If
            If rstdays(2) = True Then
                frmTrain.chkMon.Value = 1
            End If
            If rstdays(3) = True Then
                frmTrain.chkTue.Value = 1
            End If
            If rstdays(4) = True Then
                frmTrain.chkWed.Value = 1
            End If
            If rstdays(5) = True Then
                frmTrain.chkThr.Value = 1
            End If
            If rstdays(6) = True Then
                frmTrain.chkFri.Value = 1
            End If
            If rstdays(7) = True Then
                frmTrain.chkSat.Value = 1
            End If
        End If
        rstdays.Close
        rstdays.Open "select * from days where trainno= '" & downTrainNo & "' ", railCn
        If rstdays.RecordCount > 0 Then
            If rstdays(1) = True Then
                frmTrain.chkSun1.Value = 1
            End If
            If rstdays(2) = True Then
                frmTrain.chkMon1.Value = 1
            End If
            If rstdays(3) = True Then
                frmTrain.chkTue1.Value = 1
            End If
            If rstdays(4) = True Then
                frmTrain.chkWed1.Value = 1
            End If
            If rstdays(5) = True Then
                frmTrain.chkThr1.Value = 1
            End If
            If rstdays(6) = True Then
                frmTrain.chkFri1.Value = 1
            End If
            If rstdays(7) = True Then
                frmTrain.chkSat1.Value = 1
            End If
        End If
    
        Set rstRoute = New ADODB.Recordset
        rstRoute.CursorLocation = adUseClient
        rstRoute.Open "select * from routeStn,station where routeID=" & rstTrain(0) & " and routestn.stnID=station.stnID order by routestnno asc", railCn
        If rstRoute.RecordCount > 0 Then
            frmTrain.flexGridRoute.Rows = 1
            frmTrain.flexGridRoute.Cols = 7
            rstRoute.MoveFirst
            For i = 1 To rstRoute.RecordCount
                frmTrain.flexGridRoute.Rows = frmTrain.flexGridRoute.Rows + 1
                frmTrain.flexGridRoute.TextMatrix(i, 0) = frmTrain.flexGridRoute.Rows - 1
                frmTrain.flexGridRoute.TextMatrix(i, 2) = rstRoute("station.stnID")
                frmTrain.flexGridRoute.TextMatrix(i, 3) = rstRoute("stnName")
                frmTrain.flexGridRoute.TextMatrix(i, 4) = "  :  "
                frmTrain.flexGridRoute.TextMatrix(i, 5) = "  :  "
                frmTrain.flexGridRoute.Row = i
                frmTrain.flexGridRoute.Col = 1
                frmTrain.flexGridRoute.CellFontName = "Wingdings"
                frmTrain.flexGridRoute.CellFontSize = 14
                frmTrain.flexGridRoute.CellAlignment = flexAlignCenterCenter
                frmTrain.flexGridRoute.Text = strUnChecked
                rstRoute.MoveNext
            Next
            frmTrain.flexGridRoute1.Rows = 1
            frmTrain.flexGridRoute1.Cols = 7
            rstRoute.MoveLast
            For i = 1 To rstRoute.RecordCount
                frmTrain.flexGridRoute1.Rows = frmTrain.flexGridRoute1.Rows + 1
                frmTrain.flexGridRoute1.TextMatrix(i, 0) = frmTrain.flexGridRoute1.Rows - 1
                frmTrain.flexGridRoute1.TextMatrix(i, 2) = rstRoute("station.stnID")
                frmTrain.flexGridRoute1.TextMatrix(i, 3) = rstRoute("stnName")
                frmTrain.flexGridRoute1.TextMatrix(i, 4) = "  :  "
                frmTrain.flexGridRoute1.TextMatrix(i, 5) = "  :  "
                frmTrain.flexGridRoute1.Row = i
                frmTrain.flexGridRoute1.Col = 1
                frmTrain.flexGridRoute1.CellFontName = "Wingdings"
                frmTrain.flexGridRoute1.CellFontSize = 14
                frmTrain.flexGridRoute1.CellAlignment = flexAlignCenterCenter
                frmTrain.flexGridRoute1.Text = strUnChecked
                rstRoute.MovePrevious
            Next
        End If
        
        Set rstTrainRoute = New ADODB.Recordset
        rstTrainRoute.CursorLocation = adUseClient
        rstTrainRoute.Open "select * from trainroute where trainno= '" & upTrainNo & "' ", railCn
        If rstTrainRoute.RecordCount > 0 Then
            rstTrainRoute.MoveFirst
            For i = 1 To rstTrainRoute.RecordCount
                For j = 1 To frmTrain.flexGridRoute.Rows - 1
                    If frmTrain.flexGridRoute.TextMatrix(j, 2) = rstTrainRoute("stnID") Then
                        frmTrain.flexGridRoute.TextMatrix(j, 1) = strChecked
                        frmTrain.flexGridRoute.TextMatrix(j, 4) = convertTime(rstTrainRoute("arrtime")) 'rstTrainRoute("arrtime")
                        frmTrain.flexGridRoute.TextMatrix(j, 5) = convertTime(rstTrainRoute("deptime"))
                        frmTrain.flexGridRoute.TextMatrix(j, 6) = rstTrainRoute("day")
                    End If
                Next
                rstTrainRoute.MoveNext
            Next
        End If
        rstTrainRoute.Close
        rstTrainRoute.Open "select * from trainroute where trainno= '" & downTrainNo & "' ", railCn
        If rstTrainRoute.RecordCount > 0 Then
            rstTrainRoute.MoveFirst
            For i = 1 To rstTrainRoute.RecordCount
                For j = 1 To frmTrain.flexGridRoute1.Rows - 1
                    If frmTrain.flexGridRoute1.TextMatrix(j, 2) = rstTrainRoute("stnID") Then
                        frmTrain.flexGridRoute1.TextMatrix(j, 1) = strChecked
                        frmTrain.flexGridRoute1.TextMatrix(j, 4) = convertTime(rstTrainRoute("arrtime")) 'rstTrainRoute("arrtime")
                        frmTrain.flexGridRoute1.TextMatrix(j, 5) = convertTime(rstTrainRoute("deptime"))
                        frmTrain.flexGridRoute1.TextMatrix(j, 6) = rstTrainRoute("day")
                    End If
                Next
                rstTrainRoute.MoveNext
            Next
        End If
        frmTrain.txtUTrainNo.Enabled = True
        frmTrain.txtDTrainNo.Enabled = True
        frmTrain.txtTrainName.Enabled = True
        frmTrain.cmbTrainType.Enabled = True
        frmTrain.cmbRoute.Enabled = True
        frmTrain.chkMon.Enabled = True
        frmTrain.chkTue.Enabled = True
        frmTrain.chkWed.Enabled = True
        frmTrain.chkThr.Enabled = True
        frmTrain.chkFri.Enabled = True
        frmTrain.chkSat.Enabled = True
        frmTrain.chkSun.Enabled = True
        frmTrain.chkMon1.Enabled = True
        frmTrain.chkTue1.Enabled = True
        frmTrain.chkWed1.Enabled = True
        frmTrain.chkThr1.Enabled = True
        frmTrain.chkFri1.Enabled = True
        frmTrain.chkSat1.Enabled = True
        frmTrain.chkSun1.Enabled = True
        frmTrain.txtUTrainNo.BackColor = vbHighlightText
        frmTrain.txtDTrainNo.BackColor = vbHighlightText
        frmTrain.txtTrainName.BackColor = vbHighlightText
        frmTrain.cmbTrainType.BackColor = vbHighlightText
        frmTrain.cmbRoute.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmTrainDialog.Height = 1515
End Sub

Private Sub txtTrain_Change()
    Set rstTrain = New ADODB.Recordset
    rstTrain.CursorLocation = adUseClient
    rstTrain.Open "select uptrainno,trainname from train where trainname like '%" & txtTrain & "%' ", railCn
    If rstTrain.RecordCount > 0 Then
        Set DataList.RowSource = rstTrain
        Set DataList.DataSource = rstTrain
        DataList.BoundColumn = rstTrain.Fields(0).Name
        DataList.ListField = rstTrain.Fields(1).Name
        frmTrainDialog.Height = 6090
    End If
End Sub
Private Sub txtTrain_KeyPress(KeyAscii As Integer)
Call validation(2, KeyAscii, txtTrain)
End Sub

