VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form8 
   Caption         =   "Enquiry"
   ClientHeight    =   10260
   ClientLeft      =   195
   ClientTop       =   510
   ClientWidth     =   16140
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   16140
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Format          =   86900737
      CurrentDate     =   42978
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   720
      TabIndex        =   8
      Top             =   3480
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   5741
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   2
      Text            =   "To"
      Top             =   2400
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   0
      Text            =   "From"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   6120
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Continue as New passenger"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   7680
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Booking Enquiry "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination "
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
      Left            =   1440
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      Height          =   495
      Left            =   11400
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Continue as Existing  passenger"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   7680
      Width           =   4695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   735
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   7560
      Width           =   5055
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst3 As ADODB.Recordset
Dim str As String


Private Sub Combo1_Click()
  Set rst2 = New Recordset
    rst2.CursorLocation = adUseClient
    rst2.LockType = adLockOptimistic
    rst2.CursorType = adOpenDynamic
rst2.Open "select * from flightmaster where fromdestination like '" & Combo1.Text & "'", con
Combo2.Clear
For i = 0 To rst2.RecordCount - 1

 flag = 1
    For j = 0 To Combo2.ListCount - 1
        If rst2.Fields(6) = Combo2.List(j) Then
            flag = 0
            Exit For
        End If
    Next
    If flag = 1 Then
        Combo2.AddItem rst2.Fields(6)
    End If
    rst2.MoveNext

Next
rst2.Close
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

    rst1.Open "select * from flightmaster", con
  '  Set DataGrid1.DataSource = rst1
  
   
  loadvalues
End Sub


Private Sub loadvalues()
For i = 0 To rst1.RecordCount - 1

    flag = 1
    For j = 0 To Combo1.ListCount - 1
        If rst1.Fields(6) = Combo1.List(j) Then
            flag = 0
            Exit For
        End If
    Next
    If flag = 1 Then
        Combo1.AddItem rst1.Fields(6)
    End If
    rst1.MoveNext
Next
        

End Sub

Private Sub Label29_Click()
Set rst3 = New Recordset
    rst3.CursorLocation = adUseClient
    rst3.LockType = adLockOptimistic
    rst3.CursorType = adOpenDynamic



 date1 = Format(DTPicker1.Value, "dd/mm/yyyy")
 
'rst3.Open "select fid,adt,atime,ddt,dtime from schedulemaster where adt = #" & date1 & "#  and fid in(select fid from flightmaster where fromdestination like '" & Combo1.Text & "' and todestination like '" & Combo2.Text & "')", con
'rst3.Open "select fid,atime,dtime from schedulemaster where fid in(select fid from flightmaster where fromdestination like '" & Combo1.Text & "' and todestination like '" & Combo2.Text & "')", con
rst3.Open "select s.fid,f.fname,s.ddt,s.dtime,s.adt,s.atime,f.basefare,s.sno from schedulemaster as s,flightmaster as f where f.fid = s.fid  and f.fromdestination like '" & Combo1.Text & "' and f.todestination like '" & Combo2.Text & "'", con
Set DataGrid1.DataSource = rst3

'rst3.Open "select fid,adt,atime,ddt,dtime from schedulemaster where adt = #" & date1 & "#  and fid in(select fid from flightmaster where fromdestination like '" & Combo1.Text & "' and todestination like '" & Combo2.Text & "')", con
'rst3.Open "select fid,adt,atime,ddt,dtime from schedulemaster where adt = #" & date1 & "#  and fid in(select fid from flightmaster where fromdestination like '" & Combo1.Text & "' and todestination like '" & Combo2.Text & "')", con


  DataGrid1.Columns(0).Caption = "Flight Code"
  DataGrid1.Columns(1).Caption = "Airlines"
  DataGrid1.Columns(2).Caption = "Departure Date"
  DataGrid1.Columns(3).Caption = "Departure Time"
  DataGrid1.Columns(4).Caption = "Arrival Date"
  DataGrid1.Columns(5).Caption = "Arrival Time"
  DataGrid1.Columns(6).Caption = "Base Fare"
  DataGrid1.Columns(7).Caption = "Schedule No"
  
 
End Sub

Private Sub Label3_Click()
str = rst3.Fields(0)
resp = MsgBox("Are you sure you want to proceed with flight number  " & str & " ?", vbYesNo)
If resp = vbYes Then
    fno = rst3.Fields(0)
    sno = rst3.Fields(7)
    fare = rst3.Fields(6)
    Unload Me
    Form2.Show
End If
End Sub

Private Sub Label5_Click()
str = rst3.Fields(0)
resp = MsgBox("Are you sure you want to proceed with flight number  " & str & " ?", vbYesNo)
If resp = vbYes Then
    fno = rst3.Fields(0)
    sno = rst3.Fields(7)
    fare = rst3.Fields(6)
    Unload Me
    Form10.Show
End If
End Sub
