VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form blood_details 
   Caption         =   "Form2"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15705
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Show Report"
      Enabled         =   0   'False
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
      Left            =   9360
      MaskColor       =   &H8000000F&
      TabIndex        =   18
      Top             =   8400
      Width           =   2295
   End
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
      Left            =   17880
      MaskColor       =   &H8000000F&
      TabIndex        =   17
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Last Row"
      Enabled         =   0   'False
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
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   12
      Top             =   7440
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Previous Row"
      Enabled         =   0   'False
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
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   11
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next Row"
      Enabled         =   0   'False
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
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   10
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "First Row"
      Enabled         =   0   'False
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
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   9
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove Donor"
      Enabled         =   0   'False
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
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show all Donors"
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
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid details 
      Height          =   4830
      Left            =   4800
      TabIndex        =   13
      Top             =   3360
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   8520
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   16
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rubik Medium"
         Size            =   9
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rubik Medium"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc bdAdodc1 
      Height          =   330
      Left            =   1080
      Top             =   9240
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
      Caption         =   "bdAdodc1"
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "--"
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
      Left            =   7920
      TabIndex        =   16
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total number of Donors : "
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
      Left            =   4920
      TabIndex        =   15
      Top             =   8400
      Width           =   2895
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   4560
      Top             =   3120
      Width           =   15135
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Donor Details"
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
      Left            =   4920
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   960
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Table Selection Controls"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   960
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   6735
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   19335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Blood Donor Details"
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
      Width           =   6015
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
      Picture         =   "blood_details.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1275
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000040&
      Height          =   255
      Left            =   14760
      Top             =   120
      Width           =   5895
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   255
      Left            =   9360
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
      Width           =   20175
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
Attribute VB_Name = "blood_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim cn As String
Dim d As String

Private Sub Command1_Click()
cmd.ActiveConnection = con
con.CursorLocation = adUseClient
rs.Open "select * from blood_donors", con, adOpenForwardOnly, adLockOptimistic
Set details.DataSource = rs
Label14.Caption = rs.RecordCount
Command4.Enabled = True
Command2.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
DataReport3.Show
End Sub

Private Sub Command4_Click()
d = MsgBox("Are you sure you want to remove this Donor?", vbQuestion + vbYesNo, "DELETE?")
If d = vbYes Then
rs.Delete
MsgBox "Donor Removed Successfully", vbInformation
End If
End Sub

Private Sub Command5_Click()
If Command1.Enabled = False Then
con.Close
Command1.Enabled = True
Command4.Enabled = False
Command2.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
con.Open
admin_home.Show
Else
Command1.Enabled = True
Command4.Enabled = False
Command2.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
admin_home.Show
End If
End Sub

Private Sub Command6_Click()
rs.MoveFirst
End Sub

Private Sub Command7_Click()
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
ElseIf rs.BOF Then
rs.MoveFirst
End If
End Sub

Private Sub Command8_Click()
rs.MovePrevious
If rs.EOF Then
rs.MoveFirst
ElseIf rs.BOF Then
rs.MoveFirst
End If
End Sub

Private Sub Command9_Click()
rs.MoveLast
If rs.EOF Then
rs.MoveFirst
End If
End Sub

Private Sub Form_Load()
cn = "Provider=MSDAORA.1;Password=open;User ID=scott;Persist Security Info=True"
con.Open cn
End Sub

