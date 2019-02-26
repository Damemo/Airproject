VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "Schedule"
   ClientHeight    =   10260
   ClientLeft      =   195
   ClientTop       =   510
   ClientWidth     =   15735
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      DataField       =   "fid"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3120
      TabIndex        =   25
      Top             =   2040
      Width           =   2172
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   7200
      Top             =   7320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1720
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\project flight\flightmaster.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\project flight\flightmaster.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "schedulemaster"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   6840
      TabIndex        =   21
      Top             =   1920
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9
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
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "adt"
      DataSource      =   "Adodc1"
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Enter Arrival Date"
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   118816769
      CurrentDate     =   42973
   End
   Begin VB.TextBox Text1 
      DataField       =   "sno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3120
      TabIndex        =   15
      Top             =   1320
      Width           =   2172
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "atime"
      DataSource      =   "Adodc1"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   2
      ToolTipText     =   "Enter Arrival Time"
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   118816770
      CurrentDate     =   42973
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "ddt"
      DataSource      =   "Adodc1"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "Enter departure date"
      Top             =   4080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   118816769
      CurrentDate     =   42973
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "dtime"
      DataSource      =   "Adodc1"
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "Enter departure time"
      Top             =   4800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   118816770
      UpDown          =   -1  'True
      CurrentDate     =   42973
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "View Data"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   9720
      TabIndex        =   23
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   14160
      TabIndex        =   22
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Clears all fields"
      Top             =   7080
      Width           =   732
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   972
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11520
      TabIndex        =   17
      Top             =   5880
      Width           =   615
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   11280
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8400
      TabIndex        =   18
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10200
      TabIndex        =   19
      Top             =   5880
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   9960
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   5280
      TabIndex        =   8
      ToolTipText     =   "Exits from Form"
      Top             =   7080
      Width           =   612
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "Deletes the data"
      Top             =   7080
      Width           =   852
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12600
      TabIndex        =   16
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "Saves schedule details"
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Details"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   14
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Departure Time :- "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   13
      Top             =   4800
      Width           =   2412
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Departure Date :- "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   12
      Top             =   4080
      Width           =   2412
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Time :- "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   600
      TabIndex        =   11
      Top             =   3360
      Width           =   1932
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Arrival Date :- "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   10
      Top             =   2760
      Width           =   1932
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight ID "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   600
      TabIndex        =   9
      Top             =   2040
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule No  "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   12480
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   1092
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   972
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   13920
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1572
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   11040
      TabIndex        =   24
      ToolTipText     =   "Saves schedule details"
      Top             =   7920
      Width           =   615
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   10800
      Shape           =   4  'Rounded Rectangle
      Top             =   7800
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rst1 As ADODB.Recordset

Private Sub Command1_Click()
rst1.MoveFirst
Call ShowData

End Sub
Private Sub Command6_Click()

rst1.MoveNext
If rst1.EOF Then
rst1.MoveLast
End If

Call ShowData

End Sub


Private Sub Command8_Click()
rst1.MoveLast
Call ShowData

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
Set con = New Connection
    Set rst1 = New Recordset
    rst1.CursorLocation = adUseClient
    rst1.LockType = adLockOptimistic
    rst1.CursorType = adOpenDynamic


    With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\flightmaster.mdb;"
        .Open
    End With

    rst1.Open "select * from schedulemaster", con
    DataGrid1.Visible = False
    Label8.Visible = False
    Shape2.Visible = False
    Shape7.Visible = False
    Label12.Visible = False
    Label10.Visible = False
    Shape6.Visible = False
    Label16.Visible = False
    Shape8.Visible = False
    Label9.Visible = False
    Shape3.Visible = False
    Label18.Visible = False
    Shape10.Visible = False
    
    Set DataGrid1.DataSource = rst1
   DataGrid1.Columns(0).Caption = "Schedule No"
  DataGrid1.Columns(1).Caption = "Flight ID"
  DataGrid1.Columns(2).Caption = "Arrival Date"
  DataGrid1.Columns(3).Caption = "Arrival Time"
  DataGrid1.Columns(4).Caption = "Departure Date"
  DataGrid1.Columns(5).Caption = "Departure Time"
  
  

End Sub

Private Sub Label10_Click()
rst1.MoveNext
If rst1.EOF Then
rst1.MoveLast
End If

End Sub

Private Sub Label11_Click()
'For Each Control In Form4.Controls
' If TypeName(Control) = "TextBox" Then
'    If Control.Text = "" Then
'          MsgBox "Please Enter ALL Fields"
'          Exit Sub
'      End If
'End If
'Next
'
'
'For Each Control In Form4.Controls
' If TypeName(Control) = "ComboBox" Then
'    If Control.Text = "" Then
'          MsgBox "Please Enter ALL Fields"
'          Exit Sub
'      End If
'End If
'Next

resp = MsgBox("Are you sure you want to add the data?", vbYesNo)
    If resp = vbYes Then
        rst1.AddNew
     
        rst1.Fields(0) = Text1.Text
        rst1.Fields(1) = Text2.Text
        rst1.Fields(2) = DTPicker1(0).Value
        
        rst1.Fields(3) = DTPicker1(1).Value
        rst1.Fields(4) = DTPicker1(2).Value
        rst1.Fields(5) = DTPicker1(3).Value
        
     '   rst1.Fields(6) = Combo2.Text
        End If
        
        rst1.Update
        MsgBox "Flight schedule details added successfully. "
     

End Sub

Private Sub Label12_Click()
rst1.MovePrevious
If rst1.BOF Then
rst1.MoveFirst
End If


End Sub

Private Sub Label14_Click()
resp = MsgBox("Are you sure you want to delete the data?", vbYesNo)
    If resp = vbYes Then
        rst1.Delete
        End If
        
End Sub

Private Sub Label15_Click()
resp = MsgBox("Are you sure you want to exit?", vbYesNo)
    If resp = vbYes Then
Unload Me
End If
End Sub

Private Sub Label16_Click()
rst1.MoveLast

End Sub

Private Sub Label17_Click()
Text1.Text = ""
Combo1.Text = ""
Combo2.Text = ""

End Sub

Private Sub Label18_Click()
resp = MsgBox("Are you sure you want to Save the modified data?", vbYesNo)
    If resp = vbYes Then
    rst1.Update
    MsgBox "Details updated successfully", , "Confirmation "
End If
End Sub

Private Sub Label21_Click()
DataGrid1.Visible = True
Label8.Visible = True
    Shape2.Visible = True
    Shape7.Visible = True
    Label12.Visible = True
    Label10.Visible = True
    Shape6.Visible = True
    Label16.Visible = True
    Shape8.Visible = True
    Label9.Visible = True
    Shape3.Visible = True
    Label18.Visible = True
    Shape10.Visible = True
    
End Sub

Private Sub Label21_DblClick()
DataGrid1.Visible = False
    Label8.Visible = False
    Shape2.Visible = False
    Shape7.Visible = False
    Label12.Visible = False
    Label10.Visible = False
    Shape6.Visible = False
    Label16.Visible = False
    Shape8.Visible = False
    Label9.Visible = False
    Shape3.Visible = False
    Label18.Visible = False
    Shape10.Visible = False
End Sub

Private Sub Label8_Click()
rst1.MoveFirst

End Sub

Private Sub Label9_Click()
resp = MsgBox("Are you sure you want to update the data?", vbYesNo)
    If resp = vbYes Then
        DataGrid1.AllowUpdate = True
        End If
End Sub

'Private Sub Text2_Click()
'Text2.Text = Calendar1.Value
'Calendar1.Height = 2175
'Calendar1.Width = 3495
'End Sub
'
'
'
'Private Sub Text3_Click()
'Text3.Text = Calendar1.Value
'Calendar1.Height = 2175
'Calendar1.Width = 3495
'End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    
End Sub
