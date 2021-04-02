VERSION 5.00
Begin VB.Form home 
   Caption         =   "Form2"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16155
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Request Service"
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
      Left            =   4680
      TabIndex        =   15
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Yes, Register"
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
      Left            =   15600
      TabIndex        =   12
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Yes, Register"
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
      TabIndex        =   11
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Signout"
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
      Left            =   18240
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   705
      Left            =   17040
      Picture         =   "home.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   795
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   18000
      Top             =   720
      Width           =   15
   End
   Begin VB.Label uname 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000C0&
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14040
      TabIndex        =   17
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Get experts advice,tips : Find out more infomation about organ and blood donation,it's benefits."
      BeginProperty Font 
         Name            =   "Open Sans SemiBold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   9000
      TabIndex        =   16
      Top             =   1800
      Width           =   11055
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   2640
      Top             =   6840
      Width           =   4335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "We're here for you. Need help of donor to save a life find help here"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2295
      Left            =   2640
      TabIndex        =   14
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Request donor service"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2640
      TabIndex        =   13
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   13200
      Top             =   6840
      Width           =   4335
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   7920
      Top             =   6840
      Width           =   4335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " Give someone the gift of life : register as an INFINITY LIFE organ donor"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2295
      Left            =   7920
      TabIndex        =   10
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      Caption         =   "Become a organ donor"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   7920
      TabIndex        =   9
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please register today. Donating blood is easy, and saves lives"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2295
      Left            =   13200
      TabIndex        =   8
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Save  a life : give blood"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   13200
      TabIndex        =   7
      Top             =   3600
      Width           =   4335
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
      TabIndex        =   6
      Top             =   9600
      Width           =   20295
   End
   Begin VB.Image Image2 
      Height          =   8505
      Left            =   240
      Picture         =   "home.frx":9A07
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   19755
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Home"
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
      Left            =   5880
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   4920
      Top             =   1560
      Width           =   3855
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
      Picture         =   "home.frx":10B46A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Home"
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
      Left            =   -5280
      TabIndex        =   0
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000040&
      Height          =   255
      Left            =   14640
      Top             =   120
      Width           =   5775
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   255
      Left            =   8880
      Top             =   120
      Width           =   5775
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
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
login.Show
End Sub

Private Sub Command2_Click()
organ_reg.Show
End Sub

Private Sub Command3_Click()
blood_reg.Show
End Sub

Private Sub Command4_Click()
request_reg.Show
End Sub

