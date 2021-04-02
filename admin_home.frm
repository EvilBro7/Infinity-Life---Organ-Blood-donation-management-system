VERSION 5.00
Begin VB.Form admin_home 
   Caption         =   "Form2"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15765
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "Add New User"
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
      Left            =   3240
      MaskColor       =   &H8000000F&
      TabIndex        =   15
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Request Details"
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
      Left            =   16320
      TabIndex        =   14
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Organ Donor Details"
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
      Left            =   11520
      TabIndex        =   13
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Blood Donor Details"
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
      Left            =   6720
      TabIndex        =   12
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "User Details"
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
      Left            =   1080
      MaskColor       =   &H8000000F&
      TabIndex        =   11
      Top             =   7560
      Width           =   1935
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
      Left            =   18120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image Image7 
      Height          =   825
      Left            =   15240
      Picture         =   "admin_home.frx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   795
   End
   Begin VB.Shape Shape31 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   375
      Left            =   19440
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape30 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   375
      Left            =   14640
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape29 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   375
      Left            =   4920
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape28 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   375
      Left            =   9840
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "REQUESTS"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   16560
      TabIndex        =   10
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "ORGAN DONORS"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11400
      TabIndex        =   9
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "BLOOD DONORS"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6720
      TabIndex        =   8
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "USERS"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Shape Shape23 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   735
      Left            =   15840
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Shape22 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   735
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Shape21 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   735
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Shape20 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Shape Shape19 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   16200
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   975
      Left            =   15600
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Image Image5 
      Height          =   2505
      Left            =   16080
      Picture         =   "admin_home.frx":18668
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   3195
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   15600
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11400
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   975
      Left            =   10800
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Image Image4 
      Height          =   2745
      Left            =   11400
      Picture         =   "admin_home.frx":255DD
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2955
   End
   Begin VB.Shape Shape14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   10800
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   975
      Left            =   6000
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   3105
      Left            =   6720
      Picture         =   "admin_home.frx":3E2A9
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   2715
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   975
      Left            =   1080
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   3225
      Left            =   1560
      Picture         =   "admin_home.frx":57271
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   3195
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Label Label4 
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
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   4920
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000C0&
      Caption         =   "[ ADMIN ]"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   18
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   16080
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   17880
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Admin Home"
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
      Picture         =   "admin_home.frx":5AE73
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1275
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
   Begin VB.Shape Shape5 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000040&
      Height          =   255
      Left            =   14160
      Top             =   120
      Width           =   6255
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
   Begin VB.Image Image6 
      Height          =   8505
      Left            =   240
      Picture         =   "admin_home.frx":66AAA
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   19755
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   10455
      Left            =   120
      Top             =   0
      Width           =   20295
   End
End
Attribute VB_Name = "admin_home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
login.Show
End Sub

Private Sub Command2_Click()
user_details.Show
End Sub

Private Sub Command3_Click()
blood_details.Show
End Sub

Private Sub Command4_Click()
organ_details.Show
End Sub

Private Sub Command5_Click()
request_details.Show
End Sub

Private Sub Command6_Click()
create_account.Show
End Sub
