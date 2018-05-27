VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStationDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Station"
   ClientHeight    =   5700
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtStn 
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
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DGStn 
      Height          =   4215
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7435
      _Version        =   393216
      BackColor       =   255
      ForeColor       =   -2147483634
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
         Name            =   "Century Gothic"
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
      Caption         =   "Station Name:"
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
Attribute VB_Name = "frmStationDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstStn As ADODB.Recordset
Public stnId As Long
Public stnName As Variant
Public stnCode As Variant
Private Sub DGStn_dblClick()
    If DGStn.Row <> -1 Then
        i = DGStn.Row
        DGStn.RowBookmark (i)
        frmStation.txtStnCode.Text = DGStn.Columns(1)
        stnCode = DGStn.Columns(1)
        frmStation.txtStn.Text = DGStn.Columns(2)
        stnName = DGStn.Columns(2)
        stnId = DGStn.Columns(0)
        frmStation.txtPlateform.Text = DGStn.Columns(3)
        frmStation.cmbRegion.Text = rstStn("region")
        frmStation.txtStnCode.Enabled = True
        frmStation.txtStn.Enabled = True
        frmStation.cmbRegion.Enabled = True
        frmStation.txtPlateform.Enabled = True
        frmStation.txtStnCode.BackColor = vbHighlightText
        frmStation.txtStn.BackColor = vbHighlightText
        frmStation.txtPlateform.BackColor = vbHighlightText
        frmStation.cmbRegion.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmStationDialog.Height = 1515
End Sub

Private Sub txtStn_Change()
    Set rstStn = New ADODB.Recordset
    rstStn.CursorLocation = adUseClient
    rstStn.Open "select stnID,stnCode,stnName,plateforms as Plateforms, region as Region from region,station where stnName like '%" & txtStn & "%' and region.regionID=station.RegionID", railCn
    If rstStn.RecordCount > 0 Then
        Set DGStn.DataSource = rstStn
        frmStationDialog.Height = 6090
    Else
        Set DGStn.DataSource = Nothing
        frmStationDialog.Height = 1515
    End If
End Sub

Private Sub txtStn_KeyPress(KeyAscii As Integer)
Call validation(2, KeyAscii, txtStn)
End Sub
