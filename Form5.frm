VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form5 
   Caption         =   "Ticket Form"
   ClientHeight    =   10260
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   16116
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   16116
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      ToolTipText     =   "Enter Booking Date"
      Top             =   4680
      Width           =   2775
      _ExtentX        =   4890
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
      Format          =   107020289
      CurrentDate     =   42976
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "Form5.frx":3F303
      Left            =   2400
      List            =   "Form5.frx":3F343
      TabIndex        =   2
      ToolTipText     =   "Enter Seat No."
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   16
      Top             =   3360
      Width           =   1812
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   15
      Top             =   2640
      Width           =   1812
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   14
      Top             =   1920
      Width           =   1812
   End
   Begin VB.TextBox Text1 
      DataSource      =   "DataEnvironment4"
      Enabled         =   0   'False
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
      Left            =   2400
      TabIndex        =   13
      Top             =   1200
      Width           =   1812
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "Form5.frx":3F38E
      Left            =   2400
      List            =   "Form5.frx":3F39B
      TabIndex        =   4
      ToolTipText     =   "Enter food type"
      Top             =   6840
      Width           =   2775
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
      ItemData        =   "Form5.frx":3F400
      Left            =   2400
      List            =   "Form5.frx":3F40A
      TabIndex        =   1
      ToolTipText     =   "Enter Class Type"
      Top             =   5400
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   22
      ToolTipText     =   "Enter Booking Date"
      Top             =   4080
      Width           =   2775
      _ExtentX        =   4890
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
      Format          =   107020289
      CurrentDate     =   42976
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight Date"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   360
      TabIndex        =   23
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket"
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
      Height          =   372
      Left            =   9360
      TabIndex        =   18
      Top             =   8040
      Width           =   1212
   End
   Begin VB.Label Label16 
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
      Left            =   3840
      TabIndex        =   21
      Top             =   8400
      Width           =   852
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
      Left            =   5520
      TabIndex        =   20
      Top             =   8400
      Width           =   612
   End
   Begin VB.Label Label14 
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
      Left            =   840
      TabIndex        =   19
      Top             =   8400
      Width           =   612
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
      Height          =   372
      Left            =   2400
      TabIndex        =   17
      Top             =   8400
      Width           =   732
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket Details"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   22.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   12
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Food"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "No of passengers"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Date"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Flight ID "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule No  "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger ID "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   612
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1932
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket No  "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   492
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1572
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   1092
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   1092
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   972
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   612
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   8280
      Width           =   1092
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   732
      Left            =   9120
      Shape           =   4  'Rounded Rectangle
      Top             =   7920
      Width           =   1692
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rst1 As ADODB.Recordset

Private Sub Command1_Click()

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Combo4_KeyPress(KeyAscii As Integer)
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

    rst1.Open "select * from ticketmaster", con
    
     
    
  
   ' --------------- Code to generate passenger ID automatically --------------------------
   
    
    If rst1.RecordCount <> 0 Then
        rst1.MoveLast
        var = rst1.Fields(0)
    Else
       var = 0
    End If
    var = var + 1
   
    
    Text1.Text = var
     Text2.Text = pno
      Text3.Text = fno
      Text4.Text = sno
  
End Sub

Private Sub Label10_Click()
rst1.MoveFirst
End Sub

Private Sub Label11_Click()
rst1.MoveNext
If rst1.EOF Then
rst.MoveLast
End If

End Sub

Private Sub Label12_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo2.Text = ""
        
End Sub

Private Sub Label14_Click()
For Each Control In Form5.Controls
 If TypeName(Control) = "TextBox" Then
    If Control.Text = "" Then
          MsgBox "Please Enter ALL Fields"
          Exit Sub
      End If
End If
Next


For Each Control In Form5.Controls
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
      
        rst1.Fields(0) = Text1.Text
        rst1.Fields(1) = Text2.Text
        rst1.Fields(2) = Text4.Text
        rst1.Fields(3) = Text3.Text
        rst1.Fields(4) = DTPicker1.Value
        
        rst1.Fields(5) = Combo1.Text
        rst1.Fields(6) = Val(Combo3.Text)
        rst1.Fields(7) = Combo2.Text
        rst1.Fields(8) = DTPicker2.Value
        End If

        rst1.Update
        MsgBox "Booking details added successfully. "
        tno = Val(Text1.Text)
        numberofpas = Val(Combo3.Text)
        classtype = Combo1.ListIndex
        food = Combo2.ListIndex
     
        Form6.Show
     

    
End Sub

Private Sub Label15_Click()
resp = MsgBox("Are you sure you want to exit?", vbYesNo)
    If resp = vbYes Then
Unload Me
End If
End Sub

Private Sub Label16_Click()
resp = MsgBox("Are you sure you want to delete the data?", vbYesNo)
    If resp = vbYes Then
        rst1.Delete
        End If
End Sub

Private Sub Label17_Click()
resp = MsgBox("Are you sure you want to update the data?", vbYesNo)
    If resp = vbYes Then
        DataGrid1.AllowUpdate = True
        End If
End Sub

Private Sub Label18_Click()
rst1.MoveLast
End Sub

Private Sub Label19_Click()
rst1.MovePrevious
If rst1.BOF Then
rst1.MoveFirst
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
Form6.Show

End Sub

Private Sub Label22_Click()
 DataGrid1.Visible = True
 
     Label10.Visible = True
     Shape1.Visible = True
     Label19.Visible = True
     Shape9.Visible = True
     Label11.Visible = True
     Shape6.Visible = True
     Label18.Visible = True
     Shape8.Visible = True
     Label17.Visible = True
     Shape7.Visible = True
     Label20.Visible = True
     Shape10.Visible = True
     
End Sub

Private Sub Label22_DblClick()
 DataGrid1.Visible = False
     Label10.Visible = False
     Shape1.Visible = False
     Label19.Visible = False
     Shape9.Visible = False
     Label11.Visible = False
     Shape6.Visible = False
     Label18.Visible = False
     Shape8.Visible = False
     Label17.Visible = False
     Shape7.Visible = False
     Label20.Visible = False
     Shape10.Visible = False
     
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    End If
    
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
