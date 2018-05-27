VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmfareDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Fare"
   ClientHeight    =   2145
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbTrainType 
      BackColor       =   &H8000000E&
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
      ItemData        =   "frmCoachesDialog.frx":0000
      Left            =   2640
      List            =   "frmCoachesDialog.frx":0002
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox cmbCoachType 
      BackColor       =   &H8000000E&
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
      ItemData        =   "frmCoachesDialog.frx":0004
      Left            =   2640
      List            =   "frmCoachesDialog.frx":0006
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DGFare 
      Height          =   3255
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5741
      _Version        =   393216
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
      Caption         =   "Train Type:"
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
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frmfareDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCoachType As ADODB.Recordset
Dim rstTrainType As ADODB.Recordset
Public coachID As Long, typeId As Long
Private Sub cmbCoachType_Click()
    If cmbTrainType.ListIndex <> -1 And cmbCoachType.ListIndex <> -1 Then
        Set rstCoachType = New ADODB.Recordset
        rstCoachType.CursorLocation = adUseClient
        rstCoachType.Open "select coachtypename,typename,fare from fare,coach,traintype where coach.coachtypeid=fare.coachid and traintype.typeid =fare.typeid and fare.typeid=" & cmbTrainType.ItemData(cmbTrainType.ListIndex) & " and fare.coachid=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & " ", railCn
        If rstCoachType.RecordCount > 0 Then
            Set DGFare.DataSource = rstCoachType
            frmfareDialog.Height = 6360
        Else
            frmfareDialog.Height = 2580
        End If
    End If
End Sub

Private Sub cmbTrainType_Click()
    If cmbTrainType.ListIndex <> -1 And cmbCoachType.ListIndex <> -1 Then
        Set rstTrainType = New ADODB.Recordset
        rstTrainType.CursorLocation = adUseClient
        rstTrainType.Open "select coachtypename,typename,fare from fare,coach,traintype where coach.coachtypeid=fare.coachid and traintype.typeid =fare.typeid and fare.typeid=" & cmbTrainType.ItemData(cmbTrainType.ListIndex) & " and fare.coachid=" & cmbCoachType.ItemData(cmbCoachType.ListIndex) & " ", railCn
        If rstTrainType.RecordCount > 0 Then
            Set DGFare.DataSource = rstTrainType
            frmfareDialog.Height = 6360
        Else
            frmfareDialog.Height = 2580
        End If
    End If
End Sub


Private Sub DGFare_DblClick()
    If DGFare.Row <> -1 Then
        i = DGFare.Row
        DGFare.RowBookmark (i)
        frmFare.cmbTrainType.ListIndex = comboSearch(frmFare.cmbTrainType, cmbTrainType.ItemData(cmbTrainType.ListIndex))
        frmFare.cmbCoachType.ListIndex = comboSearch(frmFare.cmbCoachType, cmbCoachType.ItemData(cmbCoachType.ListIndex))
        typeId = cmbTrainType.ItemData(cmbTrainType.ListIndex)
        coachID = cmbCoachType.ItemData(cmbCoachType.ListIndex)
        frmFare.txtFare.Text = DGFare.Columns(2)
        frmFare.cmbTrainType.Enabled = True
        frmFare.cmbCoachType.Enabled = True
        frmFare.txtFare.Enabled = True
        frmFare.txtFare.BackColor = vbHighlightText
        frmFare.cmbCoachType.BackColor = vbHighlightText
        frmFare.cmbTrainType.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub

Private Sub Form_load()
    Set rstCoachType = New ADODB.Recordset
    rstCoachType.CursorLocation = adUseClient
    rstCoachType.Open "select * from Coach", railCn
    If rstCoachType.RecordCount > 0 Then
        i = 0
        rstCoachType.MoveFirst
        Do While Not rstCoachType.EOF
            cmbCoachType.AddItem rstCoachType(1)
            cmbCoachType.ItemData(i) = rstCoachType(0)
            rstCoachType.MoveNext
            i = i + 1
        Loop
    End If
    rstCoachType.Close
    
    Set rstTrainType = New ADODB.Recordset
    rstTrainType.CursorLocation = adUseClient
    rstTrainType.Open "select * from traintype", railCn
    If rstTrainType.RecordCount > 0 Then
        i = 0
        rstTrainType.MoveFirst
        Do While Not rstTrainType.EOF
            cmbTrainType.AddItem rstTrainType(1)
            cmbTrainType.ItemData(i) = rstTrainType(0)
            rstTrainType.MoveNext
            i = i + 1
        Loop
    End If
    rstTrainType.Close
    frmfareDialog.Height = 2580
End Sub
