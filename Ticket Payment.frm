VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Ticket Payment Form"
   ClientHeight    =   10260
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   17064
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   17064
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Form6.frx":381E3
      Left            =   2760
      List            =   "Form6.frx":38202
      TabIndex        =   2
      Text            =   "Bank of Baroda"
      Top             =   8040
      Width           =   3255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   6000
      TabIndex        =   28
      Top             =   2160
      Width           =   9975
      _ExtentX        =   17590
      _ExtentY        =   6583
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   9.6
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form6.frx":38292
      Left            =   2760
      List            =   "Form6.frx":3829F
      TabIndex        =   1
      Text            =   "Payment mode"
      Top             =   7440
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   2760
      TabIndex        =   4
      ToolTipText     =   "Enter Cheque No."
      Top             =   9240
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "Enter Credit No."
      Top             =   8640
      Width           =   3255
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   40
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   39
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2880
      TabIndex        =   38
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   360
      TabIndex        =   37
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   2760
      TabIndex        =   36
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2760
      TabIndex        =   35
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of passengers "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   360
      TabIndex        =   34
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2880
      TabIndex        =   33
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Class charges"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   32
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Food Charges"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   31
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2880
      TabIndex        =   30
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label14 
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
      Height          =   375
      Left            =   9840
      TabIndex        =   29
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Ticket "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   19.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8880
      TabIndex        =   27
      Top             =   9840
      Width           =   2415
   End
   Begin VB.Label Label26 
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
      Height          =   375
      Left            =   10920
      TabIndex        =   26
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label25 
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
      Height          =   375
      Left            =   9600
      TabIndex        =   25
      Top             =   6360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label24 
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
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label20 
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
      Left            =   6480
      TabIndex        =   21
      Top             =   6360
      Width           =   612
   End
   Begin VB.Label Label19 
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
      Left            =   8760
      TabIndex        =   20
      Top             =   7800
      Width           =   852
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "+ 700"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "+ 150"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Payment Details"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   22.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   5880
      TabIndex        =   16
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Cheque No."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit No"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank Name "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Fuel Tax "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Development Tax  "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "+5% "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "GST "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fare (amount)"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket No  "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule No  "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   1092
   End
   Begin VB.Label Label21 
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
      Left            =   7200
      TabIndex        =   22
      Top             =   7800
      Width           =   612
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   1092
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   1092
   End
   Begin VB.Label Label23 
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
      Left            =   10560
      TabIndex        =   23
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
      Left            =   10320
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   972
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   1572
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   10680
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   975
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   9720
      Width           =   2655
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   9720
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rst1 As ADODB.Recordset




Private Sub Combo1_Click()
If Combo1.ListIndex = 1 Then
Text3.Enabled = True
Text2.Enabled = False
Combo2.Enabled = True

End If
If Combo1.ListIndex = 2 Then
Text2.Enabled = True
Text3.Enabled = False
Combo2.Enabled = True

End If
If Combo1.ListIndex = 0 Then
Combo2.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
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

    rst1.Open "select * from ticketpaymentmaster", con
    DataGrid1.Visible = False
    Label20.Visible = False
    Shape6.Visible = False
    Label24.Visible = False
    Shape9.Visible = False
    Label25.Visible = False
    Shape1.Visible = False
    Label26.Visible = False
    Shape8.Visible = False
    
    
    
    Set DataGrid1.DataSource = rst1
    
'  DataGrid1.Columns(0).Caption = "Ticket Payment No"
'  DataGrid1.Columns(1).Caption = "Schedule No"
'  DataGrid1.Columns(2).Caption = "Ticket No"
'  DataGrid1.Columns(3).Caption = "Fare(Amount)"
'  DataGrid1.Columns(4).Caption = "GST"
'  DataGrid1.Columns(5).Caption = "Development Tax"
'  DataGrid1.Columns(6).Caption = "Fuel Tax"
'  DataGrid1.Columns(7).Caption = "Net Amount"
'  DataGrid1.Columns(8).Caption = "Modes of Payment"
'  DataGrid1.Columns(9).Caption = "Bank Name"
'  DataGrid1.Columns(10).Caption = "Credit No"
'  DataGrid1.Columns(11).Caption = "Cheque No"
'
  Label37.Caption = sno
  Label38.Caption = tno
  Label39.Caption = fare
  Label34.Caption = numberofpas
  
  If food = 0 Then
      Label18.Caption = "0"
  ElseIf food = 1 Then
      Label18.Caption = "300 "
  ElseIf food = 2 Then
      Label18.Caption = "500"
End If


If classtype = 0 Then
      Label32.Caption = "0"
  ElseIf classtype = 1 Then
      Label32.Caption = fare * 1.2
      
End If


net_amount = fare + fare * 0.05 + 150 + 700 + Val(Label18.Caption) + Val(Label32.Caption)
Label17.Caption = net_amount

total_amount = net_amount * numberofpas
Label35.Caption = total_amount
  
End Sub

Private Sub Label14_Click()
 DataGrid1.Visible = True
    Label20.Visible = True
    Shape6.Visible = True
    Label24.Visible = True
    Shape9.Visible = True
    Label25.Visible = True
    Shape1.Visible = True
    Label26.Visible = True
    Shape8.Visible = True
    
    
End Sub

Private Sub Label14_DblClick()
 DataGrid1.Visible = False
    Label20.Visible = False
    Shape6.Visible = False
    Label24.Visible = False
    Shape9.Visible = False
    Label25.Visible = False
    Shape1.Visible = False
    Label26.Visible = False
    Shape8.Visible = False
    
    
End Sub

Private Sub Label19_Click()
Text6.Text = ""
Text4.Text = ""
Text5.Text = ""

Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Label20_Click()
rst1.MoveFirst
End Sub

Private Sub Label21_Click()
'For Each Control In Form6.Controls
' If TypeName(Control) = "TextBox" Then
'    If Control.Text = "" Then
'          MsgBox "Please Enter ALL Fields"
'          Exit Sub
'      End If
'End If
'Next


For Each Control In Form6.Controls
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
       rst1.Fields(0) = Val(Label37.Caption)
        rst1.Fields(1) = Val(Label38.Caption)
        rst1.Fields(2) = Val(Label39.Caption)
        rst1.Fields(3) = Label5.Caption
        rst1.Fields(4) = Label15.Caption
        rst1.Fields(5) = Val(Label16.Caption)
        rst1.Fields(6) = Val(Label18.Caption)
        rst1.Fields(7) = Val(Label32.Caption)
        rst1.Fields(8) = Val(Label34.Caption)
        rst1.Fields(9) = Val(Label17.Caption)
        rst1.Fields(10) = Combo1.Text
        rst1.Fields(11) = Combo2.Text
        rst1.Fields(12) = Val(Text2.Text)
        rst1.Fields(13) = Text3.Text
        
        rst1.Fields(14) = Val(Label35.Caption)
        rst1.Fields(15) = "Booked"
        End If
        
        rst1.Update
        MsgBox "Flight Ticket paid successfully. "
     

    
End Sub

Private Sub Label22_Click()
resp = MsgBox("Are you sure you want to delete the data?", vbYesNo)
    If resp = vbYes Then
        rst1.Delete
        End If
End Sub

Private Sub Label23_Click()
resp = MsgBox("Are you sure you want to exit?", vbYesNo)
    If resp = vbYes Then
Unload Me
End If
End Sub

Private Sub Label24_Click()
rst1.MovePrevious
If rst1.BOF Then
rst1.MoveFirst
End If

End Sub

Private Sub Label25_Click()
rst1.MoveNext
If rst1.EOF Then
rst1.MoveLast
End If
End Sub

Private Sub Label26_Click()
rst1.MoveLast

End Sub

Private Sub Label27_Click()
resp = MsgBox("Are you sure you want to update the data?", vbYesNo)
    If resp = vbYes Then
        DataGrid1.AllowUpdate = True
        End If
End Sub

Private Sub Label28_Click()
resp = MsgBox("Are you sure you want to Save the modified data?", vbYesNo)
    If resp = vbYes Then
    rst1.Update
    MsgBox "Details updated successfully", , "Confirmation "
End If
End Sub

Private Sub Label29_Click()
DataEnvironment5.Command1 ((Val(Label38.Caption)))

DataReport5.Show
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    
End Sub
