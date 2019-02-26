VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Anand - Login "
   ClientHeight    =   12255
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   15885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   12255
   ScaleWidth      =   15885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   492
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "use 8 to 16 characters"
      Top             =   3600
      Width           =   3012
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   492
      Left            =   3960
      TabIndex        =   1
      ToolTipText     =   "enter username"
      Top             =   2640
      Width           =   3012
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   3840
      TabIndex        =   3
      Top             =   4800
      Width           =   972
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   6120
      TabIndex        =   4
      Top             =   4800
      Width           =   852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Constantia"
         Size            =   28.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   852
      Left            =   4320
      TabIndex        =   6
      Top             =   1320
      Width           =   2172
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   2160
      TabIndex        =   5
      Top             =   3600
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   2160
      TabIndex        =   0
      Top             =   2640
      Width           =   1572
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   732
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   1092
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   732
      Left            =   3600
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   1452
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rst1 As ADODB.Recordset

Dim flag As Integer




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

    rst1.Open "select * from login", con
    flag = 0
   
End Sub

Private Sub Label15_Click()
 resp = MsgBox("Are you sure you want to exit?", vbYesNo)
    If resp = vbYes Then
End
End If
End Sub

Private Sub Label4_Click()
For Each Control In Form1.Controls
 If TypeName(Control) = "TextBox" Then
    If Control.Text = "" Then
          MsgBox "please enter all fields"
          Exit Sub
      End If
End If

Next


passlen = Len(Text2)
If (passlen < 8 Or passlen > 20) Then
    MsgBox "Password must be between 8 to 20 characters"
    Text2.ForeColor = RGB(255, 0, 0)
    Exit Sub

End If





flag = 0
rst1.MoveFirst


While Not (rst1.EOF)

    If rst1.Fields(0) = Text1.Text And rst1.Fields(1) = Text2.Text Then

                flag = 1


    End If

        rst1.MoveNext
Wend

If flag = 1 Then
      MsgBox "Successful Login. "
      frmSplash.Show
Else
      MsgBox "UnSuccessful Login. "
End If

End Sub

