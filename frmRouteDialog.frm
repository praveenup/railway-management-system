VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRouteDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Route"
   ClientHeight    =   5025
   ClientLeft      =   9045
   ClientTop       =   1890
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRoute 
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
   Begin MSDataListLib.DataList dataList 
      Height          =   3120
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
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
      Caption         =   "Route Name:"
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
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmRouteDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstRoute As ADODB.Recordset
Dim rstRouteStn As ADODB.Recordset
Public routeId As Long
Public routeName As Variant
Dim destStopage As Integer 'for storing dest stopage
Option Explicit

Private Sub dataList_DblClick()
    If DataList.BoundText <> "" Then
        routeId = DataList.BoundText
        frmRoute.txtRoute.Text = ""
        frmRoute.cmbSource.Text = ""
        frmRoute.cmbDest.Text = ""
        frmRoute.cmbInterStn.Text = ""
        frmRoute.txtDistance.Text = ""
        frmRoute.txtStopage.Text = ""
        frmRoute.txtDestDistance.Text = ""
        frmRoute.flexGridRoute.Row = 0
        frmRoute.flexGridRoute.Rows = 1
        frmRoute.flexGridRoute.Cols = 4
        frmRoute.flexGridRoute.FixedCols = 1
        frmRoute.txtRoute.Text = DataList.Text
        routeName = DataList.Text
        Set rstRouteStn = New ADODB.Recordset
        rstRouteStn.CursorLocation = adUseClient
        
        rstRouteStn.Open "select * from routestn where routeid=" & routeId & "", railCn
            frmRoute.cmbSource.ListIndex = comboSearch(frmRoute.cmbSource, rstRouteStn("stnID"))
        rstRouteStn.Close
        
        rstRouteStn.Open "select max(routeStnNo) from routestn where routeid=" & routeId & "", railCn
            destStopage = rstRouteStn(0)
        rstRouteStn.Close
        
        rstRouteStn.Open "select routestn.stnID,stnName,routeStnNo,distance from station,routeStn where routeID=" & routeId & " and routeStnNo <> " & 1 & " and routeStnNo <> " & destStopage & " and routeStn.stnID=station.stnID order by routestnno asc"
        If rstRouteStn.RecordCount > 0 Then
            rstRouteStn.MoveFirst
            
            Dim i As Integer
            i = 1
            Do While Not rstRouteStn.EOF
                frmRoute.flexGridRoute.Rows = frmRoute.flexGridRoute.Rows + 1
                frmRoute.flexGridRoute.TextMatrix(i, 0) = rstRouteStn(0)
                frmRoute.flexGridRoute.TextMatrix(i, 1) = rstRouteStn(1)
                frmRoute.flexGridRoute.TextMatrix(i, 2) = rstRouteStn(2)
                frmRoute.flexGridRoute.TextMatrix(i, 3) = rstRouteStn(3)
                i = i + 1
                rstRouteStn.MoveNext
                
            Loop
        End If
        rstRouteStn.Close
        
        rstRouteStn.Open "select stnID,distance from routestn where routeid=" & routeId & " and routeStnNo=" & destStopage & "", railCn
            frmRoute.cmbDest.ListIndex = comboSearch(frmRoute.cmbDest, rstRouteStn(0))
            frmRoute.txtDestDistance.Text = rstRouteStn(1)
        rstRouteStn.Close
        
        frmRoute.txtRoute.Enabled = True
        frmRoute.cmbSource.Enabled = True
        frmRoute.cmbDest.Enabled = True
        frmRoute.cmbInterStn.Enabled = True
        frmRoute.txtDistance.Enabled = True
        frmRoute.txtStopage.Enabled = True
        frmRoute.txtDestDistance.Enabled = True
        frmRoute.txtRoute.BackColor = vbHighlightText
        frmRoute.cmbSource.BackColor = vbHighlightText
        frmRoute.cmbDest.BackColor = vbHighlightText
        frmRoute.cmbInterStn.BackColor = vbHighlightText
        frmRoute.txtDistance.BackColor = vbHighlightText
        frmRoute.txtStopage.BackColor = vbHighlightText
        frmRoute.txtDestDistance.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmRouteDialog.Height = 1600
End Sub

Private Sub txtRoute_Change()
    Set rstRoute = New ADODB.Recordset
    rstRoute.CursorLocation = adUseClient
    rstRoute.Open "select * from route where routeName like '" & "%" & txtRoute.Text & "%' ", railCn
    If rstRoute.RecordCount > 0 Then
        Set DataList.RowSource = rstRoute
        Set DataList.DataSource = rstRoute
        DataList.BoundColumn = rstRoute.Fields(0).Name
        DataList.ListField = rstRoute.Fields(1).Name
        frmRouteDialog.Height = 5500
    End If
End Sub


Private Sub txtRoute_KeyPress(KeyAscii As Integer)
Call validation(2, KeyAscii, txtRoute)
End Sub
