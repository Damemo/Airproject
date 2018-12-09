VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "Passenger Form"
   ClientHeight    =   10260
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   17064
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   17064
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
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
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "Enter Last Name"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Form2.frx":29C5A
      Left            =   2280
      List            =   "Form2.frx":29C6D
      TabIndex        =   5
      ToolTipText     =   "Enter Nationality"
      Top             =   6465
      Width           =   3375
   End
   Begin VB.TextBox Text3 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   7
      ToolTipText     =   "Enter 8-10 digits "
      Top             =   7200
      Width           =   3375
   End
   Begin VB.TextBox Text1 
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
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Enter First Name"
      Top             =   2265
      Width           =   3375
   End
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
      ItemData        =   "Form2.frx":29CA0
      Left            =   2280
      List            =   "Form2.frx":29CAA
      TabIndex        =   3
      ToolTipText     =   "Enter Gender"
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox Text4 
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
      Left            =   2280
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "Enter Address"
      Top             =   5160
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   3720
      Width           =   3375
      _ExtentX        =   5948
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   108658689
      CurrentDate     =   36526
      MaxDate         =   36526
   End
   Begin VB.Label Label23 
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
      Left            =   3360
      TabIndex        =   20
      ToolTipText     =   "Deletes the Data"
      Top             =   8400
      Width           =   855
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
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      ToolTipText     =   "Exits the Form"
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      ToolTipText     =   "Clears All Fields"
      Top             =   8400
      Width           =   735
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
      Height          =   375
      Left            =   840
      TabIndex        =   8
      ToolTipText     =   "Saves your info"
      Top             =   8400
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
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
      Left            =   2400
      TabIndex        =   17
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger Details"
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
      Left            =   4800
      TabIndex        =   16
      Top             =   240
      Width           =   5175
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
      Left            =   720
      TabIndex        =   15
      Top             =   5205
      Width           =   1095
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
      Left            =   720
      TabIndex        =   14
      Top             =   4500
      Width           =   1335
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
      Left            =   600
      TabIndex        =   13
      Top             =   7200
      Width           =   1455
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
      Left            =   600
      TabIndex        =   12
      Top             =   6495
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
      Left            =   480
      TabIndex        =   11
      Top             =   3675
      Width           =   1695
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
      Left            =   600
      TabIndex        =   10
      Top             =   2985
      Width           =   1575
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
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
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
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   975
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   615
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
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

'Private Sub Calendar1_Click()
'Text6.Text = Calendar1.Value
'Calendar1.Visible = False
'End Sub


'Private Sub Command9_Click()
'Calendar1.Visible = True
'Calendar1.Height = 2500
'Calendar1.Width = 3500
'Calendar1.Top = 3100
'Calendar1.Left = 2300
'
'End Sub



Private Sub Form_Load()
'Set Label10.DataSource = ""


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

    rst1.Open "select * from passengermaster", con
    
    
    
    
    ' --------------- Code to generate passenger ID automatically --------------------------
    rst2.Open "select * from passengermaster", con
    
    If rst2.RecordCount <> 0 Then
        rst2.MoveLast
        var = rst2.Fields(0)
    Else
       var = 0
    End If
    var = var + 1
     Label10.Caption = var
     rst2.Close
     
   
End Sub

Private Sub Label11_Click()

For Each Control In Form2.Controls
 If TypeName(Control) = "TextBox" Then
    If Control.Text = "" Then
          MsgBox "Please Enter ALL Fields"
          Exit Sub
      End If
End If
Next


For Each Control In Form2.Controls
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
       rst1.Fields(0) = var
        rst1.Fields(1) = Text1.Text
        rst1.Fields(2) = Text2.Text
        rst1.Fields(3) = DTPicker1.Value
        rst1.Fields(4) = Combo2.Text
        rst1.Fields(5) = Text4.Text
        rst1.Fields(7) = Combo1.Text
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

 
     
     
   



Private Sub Label15_Click()
rst1.MoveNext
If rst1.EOF Then
rst1.MoveLast
End If


End Sub

Private Sub Label16_Click()
rst1.MoveFirst

End Sub

Private Sub Label17_Click()

rst1.MoveLast
End Sub

Private Sub Label18_Click()
rst1.MovePrevious
If rst1.BOF Then
rst1.MoveFirst
End If


End Sub

Private Sub Label19_Click()
resp = MsgBox("Are you sure you want to exit?", vbYesNo)
    If resp = vbYes Then
Unload Me
End If
End Sub

Private Sub Label20_Click()
resp = MsgBox("Are you sure you want to Save the modified data?", vbYesNo)
    If resp = vbYes Then
    rst1.Update
    MsgBox "Details updated successfully", , "Confirmation "
End If
 
End Sub

Private Sub Label21_Click()
DataGrid1.Visible = True
Set DataGrid1.DataSource = Adodc1

    Label13.Visible = True
     Shape3.Visible = True
     Label20.Visible = True
     Shape10.Visible = True
      Label18.Visible = True
    Shape5.Visible = True
    Shape7.Visible = True
    Label16.Visible = True
     Label15.Visible = True
     Shape6.Visible = True
     Label17.Visible = True
     Shape8.Visible = True
     
     
       
'     DataGrid1.Columns(0).Caption = "Passenger ID"
'  DataGrid1.Columns(1).Caption = "First Name"
'  DataGrid1.Columns(2).Caption = "Last Name"
'  DataGrid1.Columns(3).Caption = "Date of Birth"
'  DataGrid1.Columns(4).Caption = "Address"
'  DataGrid1.Columns(5).Caption = "Nationality"
'  DataGrid1.Columns(6).Caption = "Contact No"
'  DataGrid1.Columns(7).Caption = "Gender"


Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label10.Visible = True


         Text1.Visible = True
         Text2.Visible = True
          Text3.Visible = True
          Text4.Visible = True
          
         DTPicker1.Visible = True
         Combo2.Visible = True
         Combo1.Visible = True

 Set Label10.DataSource = Adodc1
 Label10.DataField = "pid"
 
 Set Text1.DataSource = Adodc1
 Text1.DataField = "pfname"
 
 Set Text2.DataSource = Adodc1
 Text2.DataField = "plname"
 
 Set DTPicker1.DataSource = Adodc1
 DTPicker1.DataField = "dob"
 
 Set Text4.DataSource = Adodc1
 Text4.DataField = "padd"
 
 Set Combo1.DataSource = Adodc1
 Combo1.DataField = "nationality"

 Set Text3.DataSource = Adodc1
 Text3.DataField = "contactno"
 
 Set Combo2.DataSource = Adodc1
 Combo2.DataField = "gender"
  
 
 
     
End Sub

'Private Sub Label21_DblClick()
' DataGrid1.Visible = False
'     Label13.Visible = False
'     Shape3.Visible = False
'     Label20.Visible = False
'     Shape10.Visible = False
'    Label18.Visible = False
'    Shape5.Visible = False
'    Shape7.Visible = False
'    Label16.Visible = False
'     Label15.Visible = False
'     Shape6.Visible = False
'     Label17.Visible = False
'     Shape8.Visible = False
'
'End Sub

Private Sub Label22_Click()
DataReport1.Show

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

Private Sub Text6_Click()
Text6.Text = MonthView1.Value
'Calendar1.Height = 2175
'Calendar1.Width = 3495
End Sub

