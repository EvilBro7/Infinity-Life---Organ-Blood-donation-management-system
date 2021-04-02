VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form blood_reg 
   Caption         =   "Form2"
   ClientHeight    =   8805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15720
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton bcancel 
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
      Left            =   16560
      TabIndex        =   37
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton bsubmit 
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
      Left            =   14280
      TabIndex        =   36
      Top             =   7800
      Width           =   1695
   End
   Begin VB.OptionButton Option8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "O-"
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
      Left            =   6600
      MaskColor       =   &H00000080&
      TabIndex        =   34
      Top             =   7680
      Width           =   735
   End
   Begin VB.OptionButton Option7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "O+"
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
      Left            =   5640
      MaskColor       =   &H00000080&
      TabIndex        =   33
      Top             =   7680
      Width           =   735
   End
   Begin VB.OptionButton Option6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "B+"
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
      Left            =   5640
      MaskColor       =   &H00000080&
      TabIndex        =   32
      Top             =   7080
      Width           =   615
   End
   Begin VB.OptionButton Option5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "B-"
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
      Left            =   6600
      MaskColor       =   &H00000080&
      TabIndex        =   31
      Top             =   7080
      Width           =   615
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "AB-"
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
      Left            =   4680
      MaskColor       =   &H00000080&
      TabIndex        =   30
      Top             =   7680
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "AB+"
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
      Left            =   3720
      MaskColor       =   &H00000080&
      TabIndex        =   29
      Top             =   7680
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "A-"
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
      Left            =   4680
      MaskColor       =   &H00000080&
      TabIndex        =   28
      Top             =   7080
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "A+"
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
      Left            =   3720
      MaskColor       =   &H00000080&
      TabIndex        =   27
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox bemail 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "BDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9960
      TabIndex        =   26
      Top             =   7680
      Width           =   3420
   End
   Begin VB.TextBox bmobile 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "BDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9960
      MaxLength       =   10
      TabIndex        =   24
      Top             =   6960
      Width           =   3420
   End
   Begin VB.TextBox bpincode 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "BDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9960
      MaxLength       =   6
      TabIndex        =   22
      Top             =   6240
      Width           =   3420
   End
   Begin VB.TextBox bcity 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "BDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9960
      TabIndex        =   20
      Top             =   5520
      Width           =   3420
   End
   Begin VB.TextBox baddress 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "BDAdodc1"
      BeginProperty Font 
         Name            =   "Rubik"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9960
      TabIndex        =   18
      Top             =   4800
      Width           =   3420
   End
   Begin VB.TextBox bweight 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "BDAdodc1"
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
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   15
      Top             =   8400
      Width           =   3735
   End
   Begin VB.TextBox bdob 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      DataSource      =   "BDAdodc1"
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
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   12
      Top             =   6240
      Width           =   3735
   End
   Begin VB.TextBox bage 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "BDAdodc1"
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
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   10
      Top             =   5520
      Width           =   3735
   End
   Begin VB.TextBox bname 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "BDAdodc1"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   4800
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc BDAdodc1 
      Height          =   330
      Left            =   240
      Top             =   9240
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
      RecordSource    =   "BLOOD_DONORS"
      Caption         =   "BDAdodc1"
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
   Begin VB.Shape Shape10 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   14040
      Top             =   7440
      Width           =   4455
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
      Left            =   14040
      TabIndex        =   35
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
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
      Left            =   8280
      TabIndex        =   25
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
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
      Left            =   8280
      TabIndex        =   23
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
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
      Left            =   8280
      TabIndex        =   21
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
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
      Left            =   8280
      TabIndex        =   19
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
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
      Left            =   8280
      TabIndex        =   17
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   8160
      Top             =   4560
      Width           =   5415
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact details"
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
      Left            =   8280
      TabIndex        =   16
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Weight"
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
      Left            =   1680
      TabIndex        =   14
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Blood Type"
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
      Left            =   1680
      TabIndex        =   13
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date Of Birth"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Age"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   5640
      Width           =   615
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
      Left            =   1680
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   1560
      Top             =   4560
      Width           =   6015
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "About you"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1155
      Left            =   2760
      Picture         =   "blood_reg.frx":0000
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1575
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   6135
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   17895
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
      TabIndex        =   5
      Top             =   9600
      Width           =   20295
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Register your details and Become a blood donor. Most people can give blood, if the are fit and healthy."
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
      Left            =   5160
      TabIndex        =   4
      Top             =   1800
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Blood Donor Register"
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
      Width           =   6375
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
      Picture         =   "blood_reg.frx":6233
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
      Width           =   5775
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
Attribute VB_Name = "blood_reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As String
Dim con As New ADODB.Connection





Private Sub bcancel_Click()
bname.Text = ""
bage.Text = ""
bdob.Text = ""
bweight.Text = ""
baddress.Text = ""
bcity.Text = ""
bpincode.Text = ""
bmobile.Text = ""
bemail.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
home.Show
End Sub

Private Sub bsubmit_Click()

If bname.Text = "" Or bage.Text = "" Or bdob.Text = "" Or bweight.Text = "" Or baddress.Text = "" Or bcity.Text = "" Or bpincode.Text = "" Or bmobile.Text = "" Then
MsgBox "Fill all the Fields", vbInformation
ElseIf Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False And Option7.Value = False And Option8.Value = False Then
MsgBox "Select the Blood Group", vbInformation
ElseIf bage.Text < 18 Then
MsgBox "Above 18+ only", vbInformation
ElseIf bweight.Text < 50 Then
MsgBox "Weight lessthen 50lbs not elegible", vbInformation
ElseIf Len(bpincode.Text) < 5 Then
MsgBox "Enter Valid AreaPincode", vbInformation
ElseIf Len(bmobile.Text) < 10 Then
MsgBox "Enter Valid Mobile Number", vbInformation
Else
bdAdodc1.RecordSource = "select * from blood_donors"
bdAdodc1.Recordset.AddNew
bdAdodc1.Recordset(0) = bname.Text
bdAdodc1.Recordset(1) = bage.Text
bdAdodc1.Recordset(2) = bdob.Text

If Option1.Value = True Then
bdAdodc1.Recordset(3) = Option1.Caption
ElseIf Option2.Value = True Then
bdAdodc1.Recordset(3) = Option2.Caption
ElseIf Option3.Value = True Then
bdAdodc1.Recordset(3) = Option3.Caption
ElseIf Option4.Value = True Then
bdAdodc1.Recordset(3) = Option4.Caption
ElseIf Option5.Value = True Then
bdAdodc1.Recordset(3) = Option5.Caption
ElseIf Option6.Value = True Then
bdAdodc1.Recordset(3) = Option6.Caption
ElseIf Option7.Value = True Then
bdAdodc1.Recordset(3) = Option7.Caption
ElseIf Option8.Value = True Then
bdAdodc1.Recordset(3) = Option8.Caption
End If

bdAdodc1.Recordset(4) = bweight.Text
bdAdodc1.Recordset(5) = baddress.Text
bdAdodc1.Recordset(6) = bcity.Text
bdAdodc1.Recordset(7) = bpincode.Text
bdAdodc1.Recordset(8) = bmobile.Text
bdAdodc1.Recordset(9) = bemail.Text

bdAdodc1.Recordset.Update
MsgBox "Donor Details Successfully Registered", vbInformation
bname.Text = ""
bage.Text = ""
bdob.Text = ""
bweight.Text = ""
baddress.Text = ""
bcity.Text = ""
bpincode.Text = ""
bmobile.Text = ""
bemail.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
home.Show
End If
End Sub





Private Sub Form_Load()
cn = "Provider=MSDAORA.1;Password=open;User ID=scott;Persist Security Info=True"
End Sub


Private Sub bname_keypress(Ascii As Integer)
If Ascii = 13 And bname.Text <> "" Then
bage.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub

Private Sub bage_keypress(Ascii As Integer)
If Ascii = 13 And bage.Text <> "" Then
bdob.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub
Private Sub bweight_keypress(Ascii As Integer)
If Ascii = 13 And bweight.Text <> "" Then
baddress.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

Private Sub bcity_keypress(Ascii As Integer)
If Ascii = 13 And bcity.Text <> "" Then
bpincode.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub

Private Sub bpincode_keypress(Ascii As Integer)
If Ascii = 13 And bpincode.Text <> "" Then
bmobile.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

Private Sub bmobile_keypress(Ascii As Integer)
If Ascii = 13 And bmobile.Text <> "" Then
bemail.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

