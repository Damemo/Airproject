VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   Caption         =   "New Passenger Add form"
   ClientHeight    =   12480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13650
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   12480
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
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
      Left            =   2520
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Enter Address"
      Top             =   5160
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
      ItemData        =   "Form12.frx":62014
      Left            =   2520
      List            =   "Form12.frx":6201E
      TabIndex        =   4
      ToolTipText     =   "Enter Gender"
      Top             =   4440
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
      Left            =   2520
      TabIndex        =   3
      ToolTipText     =   "Enter First Name"
      Top             =   2265
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
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "Enter 8-10 digits "
      Top             =   7200
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
      ItemData        =   "Form12.frx":62030
      Left            =   2520
      List            =   "Form12.frx":62043
      TabIndex        =   1
      ToolTipText     =   "Enter Nationality"
      Top             =   6465
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
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "Enter Last Name"
      Top             =   3000
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   8280
      Top             =   9240
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
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
      Connect         =   $"Form12.frx":62076
      OLEDBString     =   $"Form12.frx":62101
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "passengermaster"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "dob"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   86769665
      CurrentDate     =   42973
      MaxDate         =   42989
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger ID "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   600
      TabIndex        =   20
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   720
      TabIndex        =   19
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   2985
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   3795
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   720
      TabIndex        =   16
      Top             =   6495
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   720
      TabIndex        =   15
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   4500
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   5205
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "New Passenger Details"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3840
      TabIndex        =   12
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      DataField       =   "pid"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
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
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Saves your info"
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "Clears All Fields"
      Top             =   8160
      Width           =   735
   End
   Begin VB.Label Label19 
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
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      ToolTipText     =   "Exits the Form"
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label Label23 
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
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Deletes the Data"
      Top             =   8160
      Width           =   855
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   8040
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   8040
      Width           =   1095
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim var As Integer



Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
var = 0
     Set con = New Connection
    Set rst1 = New Recordset
    rst1.CursorLocation = adUseClient
    rst1.LockType = adLockOptimistic
    rst1.CursorType = adOpenDynamic
    
    Set rst2 = New Recordset
    rst2.CursorLocation = adUseClient
    rst2.LockType = adLockOptimistic
    rst2.CursorType = adOpenDynamic


    With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\flightmaster.mdb;"
        .Open
    End With
End Sub

Private Sub Label11_Click()
resp = MsgBox("Are you sure you want to add the data?", vbYesNo)
    If resp = vbYes Then
        rst1.AddNew
       rst1.Fields(0) = var
        rst1.Fields(1) = Text1.Text
        rst1.Fields(2) = Text2.Text
        rst1.Fields(3) = DTPicker1.Value
        rst1.Fields(4) = Text4.Text
        rst1.Fields(5) = Combo1.Text
        rst1.Fields(7) = Combo2.Text
        rst1.Fields(6) = Val(Text3.Text)
        
      
        
        rst1.Update
        MsgBox "Passenger details added successfully. Press yes to to proceed for payment"
        pno = var
        Unload Me
        Form5.Show
      

    End If
End Sub

Private Sub Label12_Click()
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text3.Text = ""
Combo2.Text = ""
      
End Sub

Private Sub Label19_Click()
resp = MsgBox("Are you sure you want to exit?", vbYesNo)
    If resp = vbYes Then
Unload Me
End If
End Sub

Private Sub Label23_Click()
resp = MsgBox("Are you sure you want to delete the data?", vbYesNo)
    If resp = vbYes Then
rst1.Delete
End If
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
'------------------------------Validation for mobile no -------------------------------
    If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If


  '-----------------------------------------------------------------------------------

    
End Sub
