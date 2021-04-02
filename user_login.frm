VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form user_login 
   Caption         =   "Form2"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16155
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   16155
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc ulAdodc1 
      Height          =   330
      Left            =   480
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=MSDAORA.1;Password=open;User ID=scott;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=open;User ID=scott;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "scott"
      Password        =   "open"
      RecordSource    =   "select * from user_account"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Click here"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      TabIndex        =   12
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   10
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   7920
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   6480
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7920
      TabIndex        =   6
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Forgot my password"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9000
      TabIndex        =   11
      Top             =   7680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1395
      Left            =   9480
      Picture         =   "user_login.frx":0000
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1815
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   4695
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ORGAN AND BLOOD DONATION DATABASE MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Rubik Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   9600
      Width           =   20295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   30
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "ORGAN | BLOOD  TRANSPLANTATION"
      BeginProperty Font 
         Name            =   "Rubik Light"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "LIFE"
      BeginProperty Font 
         Name            =   "Anton"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "INFINITY"
      BeginProperty Font 
         Name            =   "Anton"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   240
      Picture         =   "user_login.frx":7EEB
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1275
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000040&
      Height          =   255
      Left            =   14520
      Top             =   120
      Width           =   5895
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   255
      Left            =   9120
      Top             =   120
      Width           =   5415
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   1575
      Left            =   4920
      Top             =   120
      Width           =   15495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   10455
      Left            =   120
      Top             =   120
      Width           =   20295
   End
End
Attribute VB_Name = "user_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim cn As String
Dim b As String


Private Sub Command1_Click()
If Text1.Text = "" And Text2.Text = "" Then
MsgBox "Enter valid 'username' and 'password'", vbExclamation
ElseIf Text1.Text = "" Then
MsgBox "Invalid 'Username'", vbInformation
ElseIf Text2.Text = "" Then
MsgBox "Invalid 'Password'", vbInformation
Else
cmd.ActiveConnection = con
con.CursorLocation = adUseClient
ulAdodc1.RecordSource = "select * from user_account where username='" + Text1.Text + "' and password='" + Text2.Text + "'"
ulAdodc1.Refresh
If ulAdodc1.Recordset.EOF Then
MsgBox "Invalid 'Username' and 'Password'", vbInformation
Else
MsgBox "Welcome " + Text1.Text + "", vbInformation
home.uname.Caption = "[ " + Text1.Text + " ]"
home.Show
Text1.Text = ""
Text2.Text = ""
End If
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
login.Show
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox "Please enter the 'Username' for 'password'", vbInformation
Else
a = InputBox("Enter your UserID")
End If
cmd.ActiveConnection = con
con.CursorLocation = adUseClient
ulAdodc1.RecordSource = "select * from user_account where userid='" + a + "'"
ulAdodc1.Refresh
If ulAdodc1.Recordset.EOF Then
MsgBox "Invalid UserID", vbInformation
Else
ulAdodc1.RecordSource = "select password from user_account where username='" + Text1.Text + "'"
Text2.Text = ulAdodc1.Recordset(2)
MsgBox "Password is " + Text2.Text + "", vbInformation
End If
End Sub

Private Sub Form_Load()
cn = "Provider=MSDAORA.1;Password=open;User ID=scott;Persist Security Info=True"
con.Open cn
End Sub

Private Sub Text1_keypress(Ascii As Integer)
If Ascii = 13 And Text1.Text <> "" Then
Text2.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub
