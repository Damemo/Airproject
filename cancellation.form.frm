VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7_1 
   Caption         =   "Cancellation Form"
   ClientHeight    =   10260
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   16404
   LinkTopic       =   "Form7"
   Picture         =   "cancellation.form.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   16404
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   6120
      TabIndex        =   24
      Top             =   2280
      Width           =   9855
      _ExtentX        =   17378
      _ExtentY        =   7641
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
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
   Begin VB.TextBox Text4 
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
      Left            =   8640
      TabIndex        =   23
      Top             =   1440
      Width           =   2412
   End
   Begin VB.TextBox Text7 
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
      Left            =   2880
      TabIndex        =   21
      Top             =   3120
      Width           =   1692
   End
   Begin VB.TextBox Text6 
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
      Left            =   2880
      TabIndex        =   12
      Top             =   3720
      Width           =   1692
   End
   Begin VB.TextBox Text5 
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
      Left            =   2880
      TabIndex        =   11
      Top             =   2520
      Width           =   1692
   End
   Begin VB.TextBox Text3 
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
      Left            =   2880
      TabIndex        =   10
      Top             =   1920
      Width           =   1692
   End
   Begin VB.TextBox Text2 
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
      Left            =   2880
      TabIndex        =   9
      Top             =   5520
      Width           =   1812
   End
   Begin VB.TextBox Text1 
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
      Left            =   2880
      TabIndex        =   1
      Top             =   4320
      Width           =   1692
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   3720
      Top             =   7920
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   16.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6960
      TabIndex        =   20
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Passenge Name"
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
      Height          =   975
      Left            =   600
      TabIndex        =   22
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label10 
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
      Height          =   375
      Left            =   13200
      TabIndex        =   19
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label9 
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
      Left            =   11760
      TabIndex        =   18
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label17 
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
      Left            =   10320
      TabIndex        =   17
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label18 
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
      Left            =   8400
      TabIndex        =   16
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label16 
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
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   7080
      Width           =   615
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
      Left            =   3360
      TabIndex        =   14
      Top             =   7080
      Width           =   492
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Left            =   1320
      TabIndex        =   13
      Top             =   7080
      Width           =   852
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "-40%"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Refund Amount "
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
      Left            =   600
      TabIndex        =   7
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction "
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
      Left            =   600
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Amount"
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
      Left            =   600
      TabIndex        =   5
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket No"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancellation Form"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   22.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   1200
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
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   972
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   8280
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   10080
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   13080
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   1215
   End
End
Attribute VB_Name = "Form7_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst3 As ADODB.Recordset

Private Sub Form_Load()
Set con = New Connection
    Set rst1 = New Recordset
    rst1.CursorLocation = adUseClient
    rst1.LockType = adLockOptimistic
    rst1.CursorType = adOpenDynamic

    Set rst2 = New Recordset
    rst2.CursorLocation = adUseClient
    rst2.LockType = adLockOptimistic
    rst2.CursorType = adOpenDynamic

Set rst3 = New Recordset
    rst3.CursorLocation = adUseClient
    rst3.LockType = adLockOptimistic
    rst3.CursorType = adOpenDynamic
    With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\flightmaster.mdb;"
        .Open
    End With

  '  rst1.Open "select * from cancellationmaster", con
    rst3.Open "select t.tno,t.pid,p.pfname,f.finalamt,t.fid from ticketpaymentmaster as f, passengermaster as p, ticketmaster as t where t.tno= f.tno and t.pid=p.pid and f.status like 'booked' ", con
   '' MsgBox rst3.RecordCount
    Set DataGrid1.DataSource = rst3
   
End Sub

Private Sub DataGrid1_Click()
con.Close
con.Open "Provider=microsoft.jet.OLEDB.4.0;data source=" & App.Path & "\flightmaster.mdb;"
rst3.CursorLocation = adUseClient
rst3.Open "select * from ticketpaymentmaster where tno =" & Text4.Text & "", con, adOpenDynamic, adLockOptimistic

Text1.Text = rst3.Fields(2)
Text2.Text = Val(Text1.Text) * 0.6

Text3.Text = rst3.Fields(0)
Text5.Text = rst3.Fields(1)
Text6.Text = rst3.Fields(4)
rst3.Close
rst3.CursorLocation = adUseClient
rst3.Open "select * from passengermaster where pid =" & Text4.Text & "", con, adOpenDynamic, adLockOptimistic

Text7.Text = rst3.Fields(2)

End Sub

Private Sub Label14_Click()
resp = MsgBox("Are you sure you want to cancel the ticket?", vbYesNo)
If resp = vbYes Then
        rst1.Open "select * from cancellationmaster", con
         rst1.AddNew
         rst1.Fields(0) = Val(Text3.Text)
        rst1.Fields(1) = Val(Text5.Text)
        rst1.Fields(2) = Val(Text6.Text)
     '   rst1.Fields(3) = Text7.Text
      '  rst1.Fields(4) = Text4.Text
        rst1.Fields(5) = Val(Text1.Text)
        rst1.Fields(6) = Val(Text2.Text)
       
        rst1.Update
        
         rst2.Open "select * from ticketpaymentmaster where tno= " & Val(Text3.Text), con
        If rst2.RecordCount <> 0 Then
         rst2.MoveFirst
         rst2.Fields(15) = "cancelled"
         rst2.Update
        End If
        
        
        rst3.Requery
        DataGrid1.Refresh
        MsgBox "Ticket cancelled successfully. "
        rst1.Close
        rst2.Close
        
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        
    
End If

    
End Sub

Private Sub Label15_Click()
resp = MsgBox("Are you sure you want to exit?", vbYesNo)
    If resp = vbYes Then
Unload Me
End If
End Sub

Private Sub Label7_Click()
'Dim rsSearch As New ADODB.Recordset
'rsSearch.CursorLocation = adUseClient
'SQLStr = "Select * from ticketpaymentmaster where tno like '" & Text4.Text & "%'"
'rsSearch.Open SQLStr, con, adOpenDynamic, adLockOptimistic
'
'If rsSearch.EOF = True Then
'MsgBox "No records found"
'Exit Sub
'End If
'Text3.Text = rsSearch.Fields(0).Value
'Text5.Text = rsSearch.Fields(1).Value
'Text6.Text = rsSearch.Fields(2).Value
'Text7.Text = rsSearch.Fields(3).Value
'Text4.Text = rsSearch.Fields(4).Value
'Text1.Text = rsSearch.Fields(5).Value
'Text2.Text = rsSearch.Fields(6).Value
con.Close
con.Open "Provider=microsoft.jet.OLEDB.4.0;data source=" & App.Path & "\flightmaster.mdb;"
rst1.CursorLocation = adUseClient
rst1.Open "select*from ticketpaymentmaster where tno =" & Text4.Text & "", con, adOpenDynamic, adLockOptimistic
'Set DataGrid1.DataSource = rst1
'rst1.MoveFirst
'Do While rst1.EOF = False
' If Text3.Text = rst1.Fields(0) Then
' Exit Sub
' End If
' rst1.MoveNext
' Loop
Set DataGrid1.DataSource = rst1


End Sub

