VERSION 5.00
Begin VB.Form admin_login 
   Caption         =   "Form2"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16155
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   9240
      TabIndex        =   11
      Top             =   7920
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
      Left            =   10320
      TabIndex        =   9
      Top             =   6960
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
      Left            =   8040
      TabIndex        =   8
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox apassword 
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
      Left            =   7800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   6120
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
      Left            =   8760
      TabIndex        =   10
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label5 
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
      Left            =   7800
      TabIndex        =   7
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "[ ADMIN ]"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   26.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   8640
      TabIndex        =   5
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   1395
      Left            =   9360
      Picture         =   "admin_login.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1455
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
      Width           =   20175
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   240
      Picture         =   "admin_login.frx":538E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1275
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
      Left            =   5520
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000040&
      Height          =   255
      Left            =   14160
      Top             =   120
      Width           =   6255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   255
      Left            =   9000
      Top             =   120
      Width           =   5175
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
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1815
      Left            =   8640
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   4695
      Left            =   7320
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   5415
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
Attribute VB_Name = "admin_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If apassword.Text = "" Then
MsgBox "Enter valid 'password'", vbCritical
ElseIf apassword.Text = "iladmin00" Then
MsgBox "Welcome Mr.Admin", vbInformation
apassword.Text = ""
admin_home.Show
Else
MsgBox "Invalid Password", vbCritical
apassword.Text = ""
End If
End Sub

Private Sub Command2_Click()
apassword.Text = ""
login.Show
End Sub

Private Sub Command3_Click()
a = InputBox("Enter your Tag")
If a = "iladmin00" Then
apassword.Text = a
MsgBox "Password is 'iladmin00'", vbInformation
Else
MsgBox "Incorrect", vbCritical
End If
End Sub

