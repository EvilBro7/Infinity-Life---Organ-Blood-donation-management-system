VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form organ_reg 
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
   Begin VB.OptionButton other 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "other"
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
      Left            =   6240
      MaskColor       =   &H00000000&
      TabIndex        =   37
      Top             =   7200
      Width           =   1215
   End
   Begin VB.OptionButton female 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "female"
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
      Left            =   4800
      MaskColor       =   &H00000000&
      TabIndex        =   36
      Top             =   7200
      Width           =   1215
   End
   Begin VB.OptionButton male 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "male"
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
      Left            =   3600
      MaskColor       =   &H00000000&
      TabIndex        =   35
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox oemail 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "ODAdodc1"
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
      TabIndex        =   34
      Top             =   8040
      Width           =   3420
   End
   Begin VB.CheckBox skin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "skin"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   16080
      TabIndex        =   32
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CheckBox Kidney 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "kidney"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   14520
      TabIndex        =   31
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CheckBox Liver 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "liver"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   16080
      TabIndex        =   30
      Top             =   5280
      Width           =   975
   End
   Begin VB.CheckBox Lungs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "lungs"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   14520
      TabIndex        =   29
      Top             =   5280
      Width           =   975
   End
   Begin VB.CheckBox Tooth 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "tooth"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   16080
      TabIndex        =   28
      Top             =   4680
      Width           =   975
   End
   Begin VB.CheckBox Eye 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "eye"
      BeginProperty Font 
         Name            =   "Open Sans"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   14520
      TabIndex        =   27
      Top             =   4680
      Width           =   975
   End
   Begin MSAdodcLib.Adodc ODAdodc1 
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
      RecordSource    =   "ORGAN_DONORS"
      Caption         =   "ODAdodc1"
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
   Begin VB.CommandButton ocancel 
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
      TabIndex        =   26
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton osubmit 
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
      Left            =   14520
      TabIndex        =   24
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox omobile 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "ODAdodc1"
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
      TabIndex        =   22
      Top             =   7200
      Width           =   3420
   End
   Begin VB.TextBox opincode 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "ODAdodc1"
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
      MaxLength       =   7
      TabIndex        =   20
      Top             =   6360
      Width           =   3420
   End
   Begin VB.TextBox ocity 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "ODAdodc1"
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
      Top             =   5520
      Width           =   3420
   End
   Begin VB.TextBox oaddress 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "ODAdodc1"
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
      TabIndex        =   16
      Top             =   4680
      Width           =   3420
   End
   Begin VB.TextBox odob 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "ODAdodc1"
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
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   12
      Top             =   6360
      Width           =   3735
   End
   Begin VB.TextBox oage 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "ODAdodc1"
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
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   10
      Top             =   5520
      Width           =   3735
   End
   Begin VB.TextBox oname 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "ODAdodc1"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label Label17 
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
      TabIndex        =   33
      Top             =   8160
      Width           =   735
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   14280
      Top             =   7200
      Width           =   4335
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
      Left            =   14280
      TabIndex        =   25
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   14280
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Choices"
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
      Left            =   14280
      TabIndex        =   23
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label16 
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
      TabIndex        =   21
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label15 
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
      TabIndex        =   19
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label13 
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
      TabIndex        =   17
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label12 
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
      TabIndex        =   15
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   8280
      Top             =   4320
      Width           =   5295
   End
   Begin VB.Label Label11 
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
      Left            =   8400
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gender"
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
      Width           =   975
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
      Top             =   6480
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
      Top             =   4800
      Width           =   855
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   1560
      Top             =   4320
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
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1155
      Left            =   2880
      Picture         =   "organ_reg.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1575
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   6135
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   17895
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Register your details and one day you may be able to donate your organs after death."
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
      Left            =   5280
      TabIndex        =   5
      Top             =   1800
      Width           =   9975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Organ Donor Register"
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
      TabIndex        =   4
      Top             =   720
      Width           =   6615
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
      Picture         =   "organ_reg.frx":6233
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1275
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
      TabIndex        =   0
      Top             =   9600
      Width           =   20295
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
Attribute VB_Name = "organ_reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As String
Dim con As New ADODB.Connection
Private Sub Command2_Click()
home.Show
End Sub

Private Sub Form_Load()
cn = "Provider=MSDAORA.1;Password=open;User ID=scott;Persist Security Info=True"

End Sub

Private Sub ocancel_Click()
oname.Text = ""
oage.Text = ""
odob.Text = ""
oaddress.Text = ""
ocity.Text = ""
opincode.Text = ""
omobile.Text = ""
oemail.Text = ""
male.Value = False
female.Value = False
other.Value = False
eye.Value = Unchecked
tooth.Value = Unchecked
lungs.Value = Unchecked
liver.Value = Unchecked
kidney.Value = Unchecked
skin.Value = Unchecked

home.Show
End Sub
Private Sub oage_keypress(Ascii As Integer)
If Ascii = 13 And oage.Text <> "" Then
odob.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

Private Sub ocity_keypress(Ascii As Integer)
If Ascii = 13 And ocity.Text <> "" Then
opincode.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub

Private Sub omobile_keypress(Ascii As Integer)
If Ascii = 13 And omobile.Text <> "" Then
oemail.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

Private Sub oname_keypress(Ascii As Integer)
If Ascii = 13 And oname.Text <> "" Then
oage.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub

Private Sub opincode_keypress(Ascii As Integer)
If Ascii = 13 And opincode.Text <> "" Then
omobile.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

Private Sub osubmit_Click()

If oname.Text = "" Or oage.Text = "" Or odob.Text = "" Or oaddress.Text = "" Or ocity.Text = "" Or opincode.Text = "" Then
MsgBox "Fill all the Fields", vbInformation
ElseIf eye.Value = Unchecked And tooth.Value = Unchecked And lungs.Value = Unchecked And liver.Value = Unchecked And kidney.Value = Unchecked And skin.Value = Unchecked Then
MsgBox "Select you choice", vbInformation
ElseIf oage.Text < 18 Then
MsgBox "Above 18+ only", vbInformation
ElseIf male.Value = False And female.Value = False And other.Value = False Then
MsgBox "Select the gender", vbInformation
ElseIf Len(opincode.Text) < 5 Then
MsgBox "Enter Valid AreaPincode", vbInformation
ElseIf Len(omobile.Text) < 10 Then
MsgBox "Enter Valid Mobile Number", vbInformation
ElseIf eye.Value = Checked And tooth.Value = Checked And lungs.Value = Checked And liver.Value = Checked And kidney.Value = Checked And skin.Value = Checked Then
MsgBox "choose one donate option", vbInformation
Else
ODAdodc1.RecordSource = "select * from organ_donors"
ODAdodc1.Recordset.AddNew
ODAdodc1.Recordset(0) = oname.Text
ODAdodc1.Recordset(1) = oage.Text
ODAdodc1.Recordset(2) = odob.Text

If male.Value = True Then
ODAdodc1.Recordset(3) = male.Caption
ElseIf female.Value = True Then
ODAdodc1.Recordset(3) = female.Caption
ElseIf other.Value = True Then
End If

ODAdodc1.Recordset(3) = other.Caption
ODAdodc1.Recordset(4) = oaddress.Text
ODAdodc1.Recordset(5) = ocity.Text
ODAdodc1.Recordset(6) = opincode.Text
ODAdodc1.Recordset(7) = omobile.Text
ODAdodc1.Recordset(8) = oemail.Text

If eye.Value = Checked Then
ODAdodc1.Recordset(9) = eye.Caption
ElseIf tooth.Value = Checked Then
ODAdodc1.Recordset(9) = tooth.Caption
ElseIf lungs.Value = Checked Then
ODAdodc1.Recordset(9) = lungs.Caption
ElseIf liver.Value = Checked Then
ODAdodc1.Recordset(9) = lungs.Caption
ElseIf kidney.Value = Checked Then
ODAdodc1.Recordset(9) = kidney.Caption
ElseIf skin.Value = Checked Then
ODAdodc1.Recordset(9) = skin.Caption
End If
ODAdodc1.Recordset.Update
MsgBox "Donor Details Successfully Registered", vbInformation
oname.Text = ""
oage.Text = ""
odob.Text = ""
oaddress.Text = ""
ocity.Text = ""
opincode.Text = ""
omobile.Text = ""
oemail.Text = ""
male.Value = False
female.Value = False
other.Value = False
eye.Value = Unchecked
tooth.Value = Unchecked
lungs.Value = Unchecked
liver.Value = Unchecked
kidney.Value = Unchecked
skin.Value = Unchecked
home.Show

End If
End Sub
