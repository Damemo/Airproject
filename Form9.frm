VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   10260
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   13884
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   13884
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   6480
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3196
      _ExtentY        =   1715
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=J:\project flight\flightmaster.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=J:\project flight\flightmaster.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "passengermaster"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      DataField       =   "padd"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   1800
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      ToolTipText     =   "Enter Address"
      Top             =   4800
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "gender"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form9.frx":62014
      Left            =   1800
      List            =   "Form9.frx":6201E
      TabIndex        =   5
      ToolTipText     =   "Enter Gender"
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "pfname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Enter First Name"
      Top             =   2025
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "contactno"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Enter 8-10 digits "
      Top             =   6600
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "nationality"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form9.frx":62030
      Left            =   1800
      List            =   "Form9.frx":62043
      TabIndex        =   2
      ToolTipText     =   "Enter Nationality"
      Top             =   5985
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataField       =   "plname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Enter Last Name"
      Top             =   2700
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form9.frx":62076
      Height          =   3615
      Left            =   6840
      TabIndex        =   0
      Top             =   2400
      Width           =   6495
      _ExtentX        =   11451
      _ExtentY        =   6371
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   10.8
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
      DataField       =   "dob"
      DataSource      =   "Adodc1"
      Height          =   492
      Left            =   1800
      TabIndex        =   7
      ToolTipText     =   "Enter DOB"
      Top             =   3360
      Width           =   3372
      _ExtentX        =   5948
      _ExtentY        =   868
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   107413505
      CurrentDate     =   42973
      MaxDate         =   42973
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   14640
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   1092
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   1920
      TabIndex        =   26
      Top             =   7800
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger ID "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   492
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   1572
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   2745
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   3435
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   6015
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   4140
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   5085
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "View Passenger Details "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   26.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   5040
      TabIndex        =   17
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      DataField       =   "pfname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   21.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   600
      TabIndex        =   15
      ToolTipText     =   "Saves your info"
      Top             =   7800
      Width           =   612
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Delete the data"
      Top             =   7800
      Width           =   852
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   14880
      TabIndex        =   13
      Top             =   6960
      Width           =   612
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   11880
      TabIndex        =   12
      ToolTipText     =   "Moves "
      Top             =   6960
      Width           =   612
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   8640
      TabIndex        =   11
      ToolTipText     =   "Moves to Previous Data"
      Top             =   6960
      Width           =   1212
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   10440
      TabIndex        =   10
      ToolTipText     =   "Moves to Next Data"
      Top             =   6960
      Width           =   612
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   7200
      TabIndex        =   9
      ToolTipText     =   "Moves to First Entry"
      Top             =   6960
      Width           =   612
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   4920
      TabIndex        =   8
      ToolTipText     =   "Exits the Form"
      Top             =   7800
      Width           =   612
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   1092
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   8436
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   1572
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   10236
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   1092
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   11640
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   1212
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   1092
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   972
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   1092
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'    var = 0
'     Set con = New Connection
'    Set rst1 = New Recordset
'    rst1.CursorLocation = adUseClient
'    rst1.LockType = adLockOptimistic
'    rst1.CursorType = adOpenDynamic
'
'    Set rst2 = New Recordset
'    rst2.CursorLocation = adUseClient
'    rst2.LockType = adLockOptimistic
'    rst2.CursorType = adOpenDynamic
'
'
'    With con
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\flightmaster.mdb;"
'        .Open
'    End With
'
'    rst1.Open "select pid,pfname,plname from passengermaster", con
'
'
'
'
'
'
'
'     Set DataGrid1.DataSource = rst1
'     DataGrid1.Columns(0).Caption = "Passenger ID"
'  DataGrid1.Columns(1).Caption = "First Name"
'  DataGrid1.Columns(2).Caption = "Last Name"
''  DataGrid1.Columns(3).Caption = "Date of Birth"
''  DataGrid1.Columns(4).Caption = "Address"
''  DataGrid1.Columns(5).Caption = "Nationality"
''  DataGrid1.Columns(6).Caption = "Contact No"
''  DataGrid1.Columns(7).Caption = "Gender"

     
   
End Sub

Private Sub Label11_Click()
resp = MsgBox("Are you sure you want to add the data?", vbYesNo)
    If resp = vbYes Then
        rst1.AddNew
       rst1.Fields(0) = var
        rst1.Fields(1) = Text1.Text
        rst1.Fields(2) = Text2.Text
        rst1.Fields(3) = DTPicker1.Value
        rst1.Fields(4) = Combo2.Text
        rst1.Fields(5) = Text4.Text
        rst1.Fields(7) = Combo1.Text
        rst1.Fields(6) = Val(Text3.Text)
        
      
        
        rst1.Update
        MsgBox "Passenger details added successfully."
        End If
        
End Sub

Private Sub Label13_Click()
resp = MsgBox("Are you sure you want to update the data?", vbYesNo)
    If resp = vbYes Then
            DataGrid1.AllowUpdate = True
           
        End If
End Sub

Private Sub Label14_Click()
resp = MsgBox("Are you sure you want to delete the data?", vbYesNo)
If resp = vbYes Then
    Adodc1.Recordset.Delete
End If
End Sub

Private Sub Label15_Click()
'Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Label16_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If

End Sub

Private Sub Label17_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Label18_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Label19_Click()
resp = MsgBox("Are you sure you want to exit?", vbYesNo)
    If resp = vbYes Then
Unload Me
End If
End Sub

