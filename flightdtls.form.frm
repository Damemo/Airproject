VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000010&
   Caption         =   "Flight Form"
   ClientHeight    =   10260
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   16524
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form3"
   Picture         =   "flightdtls.form.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   16524
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   2160
      Width           =   1935
      _ExtentX        =   3408
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   108265473
      CurrentDate     =   43013
   End
   Begin VB.TextBox Text1 
      DataField       =   "basefare"
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
      Height          =   405
      Left            =   3360
      TabIndex        =   28
      ToolTipText     =   "Shows Capacity"
      Top             =   6360
      Width           =   2052
   End
   Begin VB.TextBox Text3 
      DataField       =   "capacity"
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
      Height          =   405
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Shows Capacity"
      Top             =   3600
      Width           =   2052
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "fname"
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
      Height          =   372
      ItemData        =   "flightdtls.form.frx":62014
      Left            =   3360
      List            =   "flightdtls.form.frx":6201E
      TabIndex        =   2
      ToolTipText     =   "Choose Flight Name"
      Top             =   4320
      Width           =   2052
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "todestination"
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
      Height          =   372
      ItemData        =   "flightdtls.form.frx":62035
      Left            =   3360
      List            =   "flightdtls.form.frx":62054
      TabIndex        =   4
      ToolTipText     =   "Choose Destination"
      Top             =   5760
      Width           =   2052
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "fromdestination"
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
      Height          =   372
      ItemData        =   "flightdtls.form.frx":620A9
      Left            =   3360
      List            =   "flightdtls.form.frx":620C8
      TabIndex        =   3
      ToolTipText     =   "Choose Source "
      Top             =   5040
      Width           =   2052
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   7680
      Top             =   7200
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4678
      _ExtentY        =   1503
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\sem V project\flightmaster.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\sem V project\flightmaster.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3615
      Left            =   6840
      TabIndex        =   29
      Top             =   2160
      Width           =   9855
      _ExtentX        =   17378
      _ExtentY        =   6371
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   12
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
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Fare"
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
      Height          =   372
      Left            =   1200
      TabIndex        =   27
      Top             =   6360
      Width           =   1692
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      Height          =   375
      Left            =   11520
      TabIndex        =   26
      ToolTipText     =   "Moves to First Entry"
      Top             =   7200
      Width           =   855
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   11280
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "View Data"
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
      TabIndex        =   9
      ToolTipText     =   "Views DataGrid"
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Domestic"
      DataField       =   "ftype"
      DataSource      =   "Adodc1"
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
      Height          =   372
      Left            =   3360
      TabIndex        =   25
      Top             =   2880
      Width           =   1572
   End
   Begin VB.Label Label19 
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
      Left            =   15480
      TabIndex        =   15
      ToolTipText     =   "Save the update details"
      Top             =   6120
      Width           =   612
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3360
      TabIndex        =   24
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
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
      Left            =   2520
      TabIndex        =   6
      ToolTipText     =   "Clears all Fields"
      Top             =   7200
      Width           =   732
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   1092
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
      Left            =   12720
      TabIndex        =   13
      ToolTipText     =   "Moves "
      Top             =   6120
      Width           =   612
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   12480
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   972
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
      Left            =   9600
      TabIndex        =   11
      ToolTipText     =   "Moves to Previous Data"
      Top             =   6120
      Width           =   1212
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   9396
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   1572
   End
   Begin VB.Label Label12 
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
      Left            =   11400
      TabIndex        =   12
      ToolTipText     =   "Moves to Next Data"
      Top             =   6120
      Width           =   612
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   11196
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   1092
   End
   Begin VB.Label Label10 
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
      Left            =   8160
      TabIndex        =   10
      ToolTipText     =   "Moves to First Entry"
      Top             =   6120
      Width           =   612
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   1092
   End
   Begin VB.Label Label15 
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
      Left            =   6240
      TabIndex        =   8
      ToolTipText     =   "Exits the Form"
      Top             =   7200
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
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      ToolTipText     =   "Deletes the Data"
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label9 
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
      Left            =   13920
      TabIndex        =   14
      ToolTipText     =   "Update details "
      Top             =   6120
      Width           =   972
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
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "Saves Flight Details"
      Top             =   7200
      Width           =   612
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Details"
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
      Height          =   852
      Left            =   4680
      TabIndex        =   22
      Top             =   120
      Width           =   3372
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Name"
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
      Left            =   840
      TabIndex        =   21
      Top             =   4272
      Width           =   1572
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "To Destination "
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
      Height          =   732
      Left            =   840
      TabIndex        =   20
      Top             =   5760
      Width           =   2052
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "From Destination"
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
      Left            =   840
      TabIndex        =   19
      Top             =   5040
      Width           =   2292
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Capacity"
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
      Left            =   840
      TabIndex        =   18
      Top             =   3600
      Width           =   1572
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Type"
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
      Left            =   840
      TabIndex        =   17
      Top             =   2880
      Width           =   1452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Date"
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
      Left            =   840
      TabIndex        =   16
      Top             =   2160
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight ID "
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
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   1092
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   13800
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   1212
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   972
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   15240
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   1092
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   10320
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   1572
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rst1 As ADODB.Recordset


Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

'Private Sub Calendar1_Click()
'Text2.Text = Calendar1.Value
'Calendar1.Height = 100
'Calendar1.Width = 100
'End Sub



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

    rst1.Open "select * from flightmaster", con
    DataGrid1.Visible = False
    Label10.Visible = False
    Shape2.Visible = False
    Label16.Visible = False
    Shape7.Visible = False
    Label12.Visible = False
    Shape6.Visible = False
    Label17.Visible = False
    Shape8.Visible = False
   
    Shape3.Visible = False
    Label9.Visible = False
    Shape10.Visible = False
    Label19.Visible = False
    
    
    
   ' Set DataGrid1.DataSource = rst1
    
'  DataGrid1.Columns(0).Caption = "Flight ID"
'  DataGrid1.Columns(1).Caption = "Flight Date"
'  DataGrid1.Columns(2).Caption = "Flight Type"
'''  DataGrid1.Columns(3).Caption = "Capacity"
'  DataGrid1.Columns(4).Caption = "Flight Name"
'  DataGrid1.Columns(5).Caption = "From Source"
'  DataGrid1.Columns(6).Caption = "To Destination"
  

   
End Sub

Private Sub Label10_Click()
rst1.MoveFirst

End Sub

Private Sub Label11_Click()
'For Each Control In Form3.Controls
' If TypeName(Control) = "TextBox" Then
'    If Control.Text = "" Then
'          MsgBox "Please Enter ALL Fields"
'          Exit Sub
'      End If
'End If
'Next


For Each Control In Form3.Controls
 If TypeName(Control) = "ComboBox" Then
    If Control.Text = "" Then
          MsgBox "Please Enter ALL Fields"
          Exit Sub
      End If
End If
Next





resp = MsgBox("Are you sure you want to add the data?", vbYesNo)
    If resp = vbYes Then
        rst1.AddNew
      
        rst1.Fields(0) = Label8.Caption
        rst1.Fields(1) = DTPicker1.Value
        rst1.Fields(2) = Label20.Caption
        rst1.Fields(3) = Text3.Text
        rst1.Fields(4) = Combo4.Text
        rst1.Fields(5) = Combo2.Text
        rst1.Fields(6) = Combo3.Text
        rst1.Fields(7) = Text1.Text
        End If
        
        MsgBox "Flight details added successfully. "
     

End Sub


Private Sub Label12_Click()
rst1.MoveNext
If rst1.EOF Then
rst1.MoveLast
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
rst1.MovePrevious
If rst1.BOF Then
rst1.MoveFirst
End If

End Sub

Private Sub Label17_Click()
rst1.MoveLast

End Sub

Private Sub Label18_Click()
Text1.Text = ""
Text3.Text = ""

Combo4.Text = ""
Combo2.Text = ""
Combo3.Text = ""

End Sub

Private Sub Label19_Click()
resp = MsgBox("Are you sure you want to Save the modified data?", vbYesNo)
    If resp = vbYes Then
    rst1.Update
    MsgBox "Details updated successfully", , "Confirmation "
End If
 
End Sub

Private Sub Label21_Click()
DataGrid1.Visible = True
Label10.Visible = True
    Shape2.Visible = True
    Label16.Visible = True
    
    Shape7.Visible = True
    Label12.Visible = True
    Shape6.Visible = True
    Label17.Visible = True
    Shape8.Visible = True
  
    Shape3.Visible = True
    Label9.Visible = True
    Shape10.Visible = True
    Label19.Visible = True
        
End Sub

Private Sub Label21_DblClick()
DataGrid1.Visible = False
    Label10.Visible = False
    Shape2.Visible = False
    Label16.Visible = False
    Shape7.Visible = False
    Label12.Visible = False
    Shape6.Visible = False
    Label17.Visible = False
    Shape8.Visible = False
    
    Shape3.Visible = False
    Label9.Visible = False
    Shape10.Visible = False
    Label19.Visible = False
End Sub

Private Sub Label22_Click()
DataReport2.Show
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



Private Sub Text3_Change()

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    
End Sub
