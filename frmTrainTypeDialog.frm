VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTrainTypeDialog 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Train Type"
   ClientHeight    =   5775
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DGTrainType 
      Height          =   4215
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7435
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
   Begin VB.TextBox txtType 
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
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmTrainTypeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTrainType As ADODB.Recordset
Public typeId As Long
Private Sub DGTrainType_dblClick()
    If DGTrainType.Row <> -1 Then
        i = DGTrainType.Row
        DGTrainType.RowBookmark (i)
        frmTrainType.txtType.Text = DGTrainType.Columns(1)
        frmTrainType.txtSpeed.Text = DGTrainType.Columns(2)
        If DGTrainType.Columns(3) = 0 Then
            frmTrainType.opNo = True
        Else
            frmTrainType.opYes = True
        End If
        typeId = DGTrainType.Columns(0)
        frmTrainType.txtType.Enabled = True
        frmTrainType.txtSpeed.Enabled = True
        frmTrainType.opYes.Enabled = True
        frmTrainType.opNo.Enabled = True
        frmTrainType.txtType.BackColor = vbHighlightText
        frmTrainType.txtSpeed.BackColor = vbHighlightText
        saveUpdate = 2
        Unload Me
    End If
End Sub


Private Sub Form_load()
    frmTrainTypeDialog.Height = 1515
End Sub

Private Sub txtType_Change()
    Set rstTrainType = New ADODB.Recordset
    rstTrainType.CursorLocation = adUseClient
    rstTrainType.Open "select typeID as TypeID,typeName as TypeName,maxSpeed as Speed,catering as Catering from trainType where typeName like '%" & txtType & "%' ", railCn
    If rstTrainType.RecordCount > 0 Then
        Set DGTrainType.DataSource = rstTrainType
        frmTrainTypeDialog.Height = 6130
    Else
        frmTrainTypeDialog.Height = 1515
    End If
End Sub




Private Sub txtType_KeyPress(KeyAscii As Integer)
Call validation(2, KeyAscii, txtType)
End Sub
