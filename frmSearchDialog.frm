VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSearchDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Train"
   ClientHeight    =   5670
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   3720
      TabIndex        =   9
      Top             =   1920
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   8400
      TabIndex        =   8
      Top             =   1920
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000C0&
      ForeColor       =   &H00C0C000&
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtDest 
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
      Left            =   8520
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtSource 
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
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdFind 
      Height          =   375
      Left            =   5520
      Picture         =   "frmSearchDialog.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid flexGridTrain 
      CausesValidation=   0   'False
      Height          =   3975
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7011
      _Version        =   393216
      BackColor       =   -2147483634
      ForeColor       =   192
      BackColorSel    =   192
      GridColor       =   255
      GridColorFixed  =   192
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000C0&
      Height          =   4215
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Trains List"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   13215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Destination Station:"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Source Station:"
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
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmSearchDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstSourceStn As ADODB.Recordset
Dim rstDestStn As ADODB.Recordset
Dim rstTrain As ADODB.Recordset
Dim rstTrainRoute As ADODB.Recordset
Dim rstdays As ADODB.Recordset
Dim paraTrain1 As ADODB.Parameter
Dim paraTrain As ADODB.Parameter
Dim cmdTrain As ADODB.Command
Dim sourceID As Integer
Dim destID As Integer
Dim flag As Integer
Private Sub unloadOption()
If flag <> 0 Then
    For i = 1 To flag
        Unload Option1(i)
    Next
End If
End Sub
Private Sub cmdFind_Click()
If txtSource.Text <> "" And txtDest.Text <> "" Then
        
        Set cmdTrain = New ADODB.Command
        cmdTrain.CommandType = adCmdTable
        cmdTrain.CommandText = "query3"
        cmdTrain.ActiveConnection = railCn
        Set paraTrain = cmdTrain.CreateParameter("source", adInteger, adParamInput)
        cmdTrain.Parameters.Append paraTrain
        paraTrain.Value = sourceID
        Set paraTrain1 = cmdTrain.CreateParameter("dest", adInteger, adParamInput)
        cmdTrain.Parameters.Append paraTrain1
        paraTrain1.Value = destID
        Set rstTrain = cmdTrain.Execute
        If rstTrain.RecordCount > 0 Then
            unloadOption
            flexGridTrain.Rows = 1
            flexGridTrain.Cols = 6
            flexGridTrain.TextMatrix(0, 3) = "Departure From " & txtSource.Text
            flexGridTrain.TextMatrix(0, 4) = "Arrival To " & txtDest.Text
            rstTrain.MoveFirst
            For i = 1 To rstTrain.RecordCount
                    flag = i
                    flexGridTrain.Rows = flexGridTrain.Rows + 1
                    flexGridTrain.TextMatrix(i, 0) = i
                    flexGridTrain.TextMatrix(i, 1) = rstTrain("trainno")
                    flexGridTrain.TextMatrix(i, 2) = rstTrain("trainname")
                    flexGridTrain.TextMatrix(i, 3) = convertTime(rstTrain("deptime"))
                    flexGridTrain.TextMatrix(i, 4) = convertTime(rstTrain("arrtime"))
                    Load Option1(i)
                    'Option1(i).Move flexGridTrain.CellLeft + flexGridTrain.Left, flexGridTrain.Top + flexGridTrain.CellTop, flexGridTrain.CellWidth, flexGridTrain.CellHeight
                    Option1(i).Move Option1(i - 1).Left, Option1(i - 1).Top + Option1(i - 1).Height, flexGridTrain.CellWidth, flexGridTrain.CellHeight
                    Option1(i).Visible = True
                    Set rstdays = New ADODB.Recordset
                    rstdays.CursorLocation = adUseClient
                    rstdays.Open "select *from days where trainno='" & flexGridTrain.TextMatrix(i, 1) & "' ", railCn
                    rstdays.MoveFirst
                    For j = 1 To 7
                        If rstdays(j) = True Then
                            
                            If flexGridTrain.TextMatrix(i, 5) = "" Then
                                flexGridTrain.TextMatrix(i, 5) = "|" & rstdays(j).Name & "|"
                            Else
                                flexGridTrain.TextMatrix(i, 5) = flexGridTrain.TextMatrix(i, 5) & rstdays(j).Name & "|"
                            End If
                        End If
                    Next
                    rstTrain.MoveNext
            Next
            Label4.Caption = "Trains List From " & txtSource.Text & " To " & txtDest.Text
        Else
            MsgBox "Trains B/w Stn Not Found", vbCritical
        End If
Else
    MsgBox "Please Select The Source And Destination Station", vbCritical
End If
End Sub

Private Sub flexGridTrain_Click()
    If flexGridTrain.Rows > 1 Then
        frmSeatAvailablity.txtTrainNo = flexGridTrain.TextMatrix(flexGridTrain.Row, 1)
        Option1(flexGridTrain.Row).Value = True
        
        Unload Me
    End If
End Sub

Private Sub Form_load()
    flag = 0
    Option1(0).Visible = False
    List1.Visible = False
    List1.Top = txtSource.Top + txtSource.Height
    List1.Left = txtSource.Left
    List2.Visible = False
    List2.Left = txtDest.Left
    List2.Top = txtDest.Top + txtDest.Height
    flexGridTrain.Rows = 1
    flexGridTrain.Cols = 6
    flexGridTrain.TextMatrix(0, 0) = "S.No."
    flexGridTrain.TextMatrix(0, 1) = "Train No."
    flexGridTrain.TextMatrix(0, 2) = "Train Name"
    flexGridTrain.TextMatrix(0, 3) = "Departure"
    flexGridTrain.TextMatrix(0, 4) = "Arrival"
    flexGridTrain.TextMatrix(0, 5) = "Runs On"
    flexGridTrain.ColWidth(0) = 600
    flexGridTrain.ColWidth(1) = 1000
    flexGridTrain.ColWidth(2) = 2500
    flexGridTrain.ColWidth(3) = 3500
    flexGridTrain.ColWidth(4) = 3500
    flexGridTrain.ColWidth(5) = 3650
End Sub

Private Sub List1_Click()
sourceID = List1.ItemData(List1.ListIndex)
txtSource = List1.Text
frmSeatAvailablity.sourceID = sourceID
frmSeatAvailablity.txtSource = txtSource
List1.Visible = False
End Sub

Private Sub Option1_Click(Index As Integer)
    If flexGridTrain.Rows > 1 Then
        frmSeatAvailablity.txtTrainNo = flexGridTrain.TextMatrix(Index, 1)
        Unload Me
    End If
End Sub

Private Sub txtsource_LostFocus()
List1.Visible = False
End Sub

Private Sub txtdest_LostFocus()
List2.Visible = False
End Sub

Private Sub txtDest_Change()
    If txtDest.Text <> "" Then
        List2.Visible = True
        List2.Clear
        Set rstDestStn = New ADODB.Recordset
        rstDestStn.CursorLocation = adUseClient
        rstDestStn.Open "select * from station where stnname like '" & "%" & txtDest.Text & "%'", railCn
        If rstDestStn.RecordCount > 0 Then
            i = 0
            rstDestStn.MoveFirst
            Do While Not rstDestStn.EOF
                List2.AddItem rstDestStn(2) & "(" & rstDestStn(1) & ")"
                List2.ItemData(i) = rstDestStn(0)
                rstDestStn.MoveNext
                i = i + 1
            Loop
        Else
            List2.Clear
            List2.Visible = False
        End If
        rstDestStn.Close
    Else
        List2.Visible = False
    End If
    
    'destID = 0
End Sub

Private Sub list2_Click()
destID = List2.ItemData(List2.ListIndex)
txtDest = List2.Text
frmSeatAvailablity.destID = destID
frmSeatAvailablity.txtDest = txtDest
List2.Visible = False
End Sub

Private Sub txtsource_Change()
    If txtSource.Text <> "" Then
        List1.Visible = True
        List1.Clear
        Set rstSourceStn = New ADODB.Recordset
        rstSourceStn.CursorLocation = adUseClient
        rstSourceStn.Open "select * from station where stnname like '" & "%" & txtSource.Text & "%'", railCn
        If rstSourceStn.RecordCount > 0 Then
            i = 0
            rstSourceStn.MoveFirst
            Do While Not rstSourceStn.EOF
                List1.AddItem rstSourceStn(2) & "(" & rstSourceStn(1) & ")"
                List1.ItemData(i) = rstSourceStn(0)
                rstSourceStn.MoveNext
                i = i + 1
            Loop
        Else
            List1.Clear
            List1.Visible = False
        End If
        rstSourceStn.Close
    Else
        List1.Visible = False
    End If
    'sourceID = 0
End Sub


