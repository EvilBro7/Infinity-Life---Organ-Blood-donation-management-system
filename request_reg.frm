VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form request_reg 
   Caption         =   "Form2"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15540
   LinkTopic       =   "Form2"
   ScaleHeight     =   8745
   ScaleWidth      =   15540
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "AB-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   15600
      MaskColor       =   &H00000080&
      TabIndex        =   40
      Top             =   6960
      Width           =   855
   End
   Begin VB.OptionButton Option7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "AB+"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   14640
      MaskColor       =   &H00000080&
      TabIndex        =   39
      Top             =   6960
      Width           =   855
   End
   Begin VB.OptionButton Option6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "O-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   16440
      MaskColor       =   &H00000080&
      TabIndex        =   38
      Top             =   6480
      Width           =   735
   End
   Begin VB.OptionButton Option5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "O+"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   15600
      MaskColor       =   &H00000080&
      TabIndex        =   37
      Top             =   6480
      Width           =   735
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "B-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   14640
      MaskColor       =   &H00000080&
      TabIndex        =   36
      Top             =   6480
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "B+"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   16440
      MaskColor       =   &H00000080&
      TabIndex        =   35
      Top             =   6000
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "A-"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   15600
      MaskColor       =   &H00000080&
      TabIndex        =   34
      Top             =   6000
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "A+"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   14640
      MaskColor       =   &H00000080&
      TabIndex        =   33
      Top             =   6000
      Width           =   615
   End
   Begin VB.OptionButton skin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "skin"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12360
      MaskColor       =   &H00000000&
      TabIndex        =   32
      Top             =   6960
      Width           =   1935
   End
   Begin VB.OptionButton lungs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lungs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   11040
      MaskColor       =   &H00000000&
      TabIndex        =   31
      Top             =   6960
      Width           =   975
   End
   Begin VB.OptionButton liver 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "liver"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12360
      MaskColor       =   &H00000000&
      TabIndex        =   30
      Top             =   6480
      Width           =   975
   End
   Begin VB.OptionButton kidney 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "kidney"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   11040
      MaskColor       =   &H00000000&
      TabIndex        =   29
      Top             =   6480
      Width           =   1095
   End
   Begin VB.OptionButton tooth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "tooth"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   12360
      MaskColor       =   &H00000000&
      TabIndex        =   28
      Top             =   6000
      Width           =   975
   End
   Begin VB.OptionButton eye 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "eye"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   11040
      MaskColor       =   &H00000000&
      TabIndex        =   27
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton rcancel 
      Caption         =   "Cancel"
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
      Left            =   13560
      TabIndex        =   25
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton rsubmit 
      Caption         =   "Submit"
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
      Left            =   11280
      TabIndex        =   24
      Top             =   8400
      Width           =   1815
   End
   Begin VB.CommandButton rorgan 
      Caption         =   "Organ"
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
      Left            =   11640
      TabIndex        =   20
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton rblood 
      Caption         =   "Blood"
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
      Left            =   14640
      TabIndex        =   19
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox remail 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "RDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5880
      TabIndex        =   17
      Top             =   8280
      Width           =   3735
   End
   Begin VB.TextBox rmobile 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "RDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5880
      MaxLength       =   10
      TabIndex        =   15
      Top             =   7560
      Width           =   3735
   End
   Begin VB.TextBox rpincode 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "RDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5880
      MaxLength       =   6
      TabIndex        =   13
      Top             =   6840
      Width           =   3735
   End
   Begin VB.TextBox rcity 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "RDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5880
      TabIndex        =   11
      Top             =   6120
      Width           =   3735
   End
   Begin VB.TextBox raddress 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "RDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5880
      TabIndex        =   9
      Top             =   5400
      Width           =   3735
   End
   Begin VB.TextBox rname 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "RDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5880
      TabIndex        =   7
      Top             =   4680
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc RDAdodc1 
      Height          =   330
      Left            =   360
      Top             =   9120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "REQUEST_DETAILS"
      Caption         =   "RDAdodc1"
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
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "We're here for you, Need help to save a life."
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
      Left            =   6120
      TabIndex        =   26
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   10920
      Top             =   8040
      Width           =   6255
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirmation"
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
      Left            =   11040
      TabIndex        =   23
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Blood group"
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
      Left            =   14640
      TabIndex        =   22
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Organ"
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
      Left            =   11040
      TabIndex        =   21
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   615
      Left            =   11040
      Top             =   4680
      Width           =   6015
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   10920
      Top             =   4440
      Width           =   6255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Request"
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
      Left            =   11040
      TabIndex        =   18
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email"
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
      Left            =   3840
      TabIndex        =   16
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mobile"
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
      Left            =   3840
      TabIndex        =   14
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pincode"
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
      Left            =   3840
      TabIndex        =   12
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "City"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   6240
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name"
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
      Left            =   3840
      TabIndex        =   6
      Top             =   4800
      Width           =   855
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   3720
      Top             =   4440
      Width           =   6015
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your details"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1155
      Left            =   4680
      Picture         =   "request_reg.frx":0000
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1575
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   6135
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   14535
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Request Service"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   5175
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
      Picture         =   "request_reg.frx":6233
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1275
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000040&
      Height          =   255
      Left            =   14640
      Top             =   120
      Width           =   5655
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   255
      Left            =   9240
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
      Width           =   15375
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
Attribute VB_Name = "request_reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As String
Dim con As New ADODB.Connection

Private Sub Form_Load()
cn = "Provider=MSDAORA.1;Password=open;User ID=scott;Persist Security Info=True"
End Sub

Private Sub rcity_keypress(Ascii As Integer)
If Ascii = 13 And rcity.Text <> "" Then
rpincode.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub

Private Sub rmobile_keypress(Ascii As Integer)
If Ascii = 13 And rmobile.Text <> "" Then
remail.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

Private Sub rpincode_keypress(Ascii As Integer)
If Ascii = 13 And rpincode.Text <> "" Then
rmobile.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

Private Sub rname_keypress(Ascii As Integer)
If Ascii = 13 And rname.Text <> "" Then
raddress.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub

Private Sub rorgan_Click()
eye.Enabled = True
tooth.Enabled = True
lungs.Enabled = True
liver.Enabled = True
kidney.Enabled = True
skin.Enabled = True
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
End Sub
Private Sub rblood_Click()
eye.Enabled = False
tooth.Enabled = False
lungs.Enabled = False
liver.Enabled = False
kidney.Enabled = False
skin.Enabled = False
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Option4.Enabled = True
Option5.Enabled = True
Option6.Enabled = True
Option7.Enabled = True
Option8.Enabled = True
eye.Value = False
tooth.Value = False
lungs.Value = False
liver.Value = False
kidney.Value = False
skin.Value = False
End Sub
Private Sub rcancel_Click()
rname.Text = ""
raddress.Text = ""
rcity.Text = ""
rpincode = ""
rmobile.Text = ""
remail.Text = ""
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
eye.Enabled = False
tooth.Enabled = False
lungs.Enabled = False
liver.Enabled = False
kidney.Enabled = False
skin.Enabled = False
eye.Value = False
tooth.Value = False
lungs.Value = False
liver.Value = False
kidney.Value = False
skin.Value = False
home.Show
End Sub

Private Sub rsubmit_Click()
If rname.Text = "" Or raddress.Text = "" Or rcity.Text = "" Or rpincode = "" Or rmobile.Text = "" Then
MsgBox "Fill all the Fields", vbInformation
ElseIf eye.Enabled = False And tooth.Enabled = False And lungs.Enabled = False And liver.Enabled = False And kidney.Enabled = False And skin.Enabled = False And Option1.Enabled = False And Option2.Enabled = False And Option3.Enabled = False And Option4.Enabled = False And Option5.Enabled = False And Option6.Enabled = False And Option7.Enabled = False And Option8.Enabled = False Then
MsgBox "Select an option", vbInformation
ElseIf eye.Value = False And tooth.Value = False And lungs.Value = False And liver.Value = False And kidney.Value = False And skin.Value = False And Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False And Option7.Value = False And Option8.Value = False Then
MsgBox "Select an option", vbInformation
ElseIf Len(rpincode.Text) < 5 Then
MsgBox "Enter Valid AreaPincode", vbInformation
ElseIf Len(rmobile.Text) < 10 Then
MsgBox "Enter Valid Mobile Number", vbInformation
Else
rdAdodc1.RecordSource = "select * from request_details"
rdAdodc1.Recordset.AddNew
rdAdodc1.Recordset(0) = rname.Text
rdAdodc1.Recordset(1) = raddress.Text
rdAdodc1.Recordset(2) = rcity.Text
rdAdodc1.Recordset(3) = rpincode.Text
rdAdodc1.Recordset(4) = rmobile.Text
rdAdodc1.Recordset(5) = remail.Text
If eye.Value = True Then
rdAdodc1.Recordset(6) = eye.Caption
ElseIf tooth.Value = True Then
rdAdodc1.Recordset(6) = tooth.Caption
ElseIf lungs.Value = True Then
rdAdodc1.Recordset(6) = lungs.Caption
ElseIf liver.Value = True Then
rdAdodc1.Recordset(6) = liver.Caption
ElseIf kidney.Value = True Then
rdAdodc1.Recordset(6) = kidney.Caption
ElseIf skin.Value = True Then
rdAdodc1.Recordset(6) = skin.Caption
ElseIf Option1.Value = True Then
rdAdodc1.Recordset(7) = Option1.Caption
ElseIf Option2.Value = True Then
rdAdodc1.Recordset(7) = Option2.Caption
ElseIf Option3.Value = True Then
rdAdodc1.Recordset(7) = Option3.Caption
ElseIf Option4.Value = True Then
rdAdodc1.Recordset(7) = Option4.Caption
ElseIf Option5.Value = True Then
rdAdodc1.Recordset(7) = Option5.Caption
ElseIf Option6.Value = True Then
rdAdodc1.Recordset(7) = Option6.Caption
ElseIf Option7.Value = True Then
rdAdodc1.Recordset(7) = Option7.Caption
ElseIf Option8.Value = True Then
rdAdodc1.Recordset(7) = Option8.Caption
End If
rdAdodc1.Recordset.Update
MsgBox "Request Successfully Registered", vbInformation
rname.Text = ""
raddress.Text = ""
rcity.Text = ""
rpincode = ""
rmobile.Text = ""
remail.Text = ""
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Option4.Enabled = False
Option5.Enabled = False
Option6.Enabled = False
Option7.Enabled = False
Option8.Enabled = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
eye.Enabled = False
tooth.Enabled = False
lungs.Enabled = False
liver.Enabled = False
kidney.Enabled = False
skin.Enabled = False
eye.Value = False
tooth.Value = False
lungs.Value = False
liver.Value = False
kidney.Value = False
skin.Value = False
home.Show
End If

End Sub
