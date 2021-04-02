VERSION 5.00
Begin VB.Form login 
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
   Begin VB.CommandButton Command5 
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
      Left            =   17760
      MaskColor       =   &H8000000F&
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADMIN"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   5
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   120
      Top             =   120
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
   Begin VB.Shape Shape9 
      Height          =   15
      Left            =   11160
      Top             =   6840
      Width           =   4095
   End
   Begin VB.Shape Shape8 
      Height          =   15
      Left            =   11160
      Top             =   4560
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   1395
      Left            =   11280
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   1755
      Left            =   11160
      Picture         =   "Form2.frx":538E
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   4935
      Left            =   10680
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "ORGAN | BLOOD  TRANSPLANTATION"
      BeginProperty Font 
         Name            =   "Rubik Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   7200
      Width           =   4815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "INFINITY"
      BeginProperty Font 
         Name            =   "Anton"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   6000
      TabIndex        =   2
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "LIFE"
      BeginProperty Font 
         Name            =   "Anton"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   8400
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   6240
      Picture         =   "Form2.frx":8F90
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   3075
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Login as"
      BeginProperty Font 
         Name            =   "Rubik Medium"
         Size            =   33
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   5775
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   10935
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
      Left            =   8520
      Top             =   120
      Width           =   6135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   20295
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
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
user_login.Show
End Sub

Private Sub Command2_Click()
admin_login.Show
End Sub

Private Sub Command3_Click()
create_account.Show
End Sub

Private Sub Command5_Click()
End
End Sub
