VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form create_account 
   Caption         =   "Form2"
   ClientHeight    =   8925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15915
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   15915
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
      Left            =   8760
      MaskColor       =   &H00000000&
      TabIndex        =   26
      Top             =   7920
      Width           =   1095
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
      Left            =   7080
      MaskColor       =   &H00000000&
      TabIndex        =   25
      Top             =   7920
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
      Left            =   5640
      MaskColor       =   &H00000000&
      TabIndex        =   24
      Top             =   7920
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   9120
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
      RecordSource    =   "USER_ACCOUNT"
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
   Begin VB.CommandButton Command2 
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
      Left            =   15240
      TabIndex        =   23
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Left            =   15240
      TabIndex        =   22
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Register User under INFINITY LIFE terms and conditions"
      BeginProperty Font 
         Name            =   "Open Sans SemiBold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10920
      TabIndex        =   21
      Top             =   7800
      Width           =   3735
   End
   Begin VB.TextBox mobiletxt 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
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
      Left            =   12480
      MaxLength       =   10
      TabIndex        =   20
      Top             =   6960
      Width           =   4455
   End
   Begin VB.TextBox emailtxt 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
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
      Left            =   12480
      TabIndex        =   19
      Top             =   6120
      Width           =   4455
   End
   Begin VB.TextBox addresstxt 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
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
      Left            =   12480
      TabIndex        =   18
      Top             =   5280
      Width           =   4455
   End
   Begin VB.TextBox passwordtxt 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
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
      IMEMode         =   3  'DISABLE
      Left            =   5520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   6960
      Width           =   4455
   End
   Begin VB.TextBox useridtxt 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
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
      Left            =   5520
      MaxLength       =   6
      TabIndex        =   11
      Top             =   6120
      Width           =   4455
   End
   Begin VB.TextBox nametxt 
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
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
      Left            =   5520
      TabIndex        =   8
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   5760
      Top             =   7440
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label15 
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
      Left            =   10920
      TabIndex        =   17
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label13 
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
      Left            =   10920
      TabIndex        =   16
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Adderss"
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
      Left            =   10920
      TabIndex        =   15
      Top             =   5400
      Width           =   975
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   10800
      Top             =   4920
      Width           =   6255
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
      Left            =   10920
      TabIndex        =   14
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   15
      Left            =   3480
      Top             =   4920
      Width           =   6615
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
      Left            =   3600
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
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
      Left            =   3600
      TabIndex        =   10
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
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
      Left            =   3600
      TabIndex        =   9
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User ID"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   6240
      Width           =   975
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
      Left            =   3600
      TabIndex        =   6
      Top             =   5400
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   1155
      Left            =   4560
      Picture         =   "create_account.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Register your details, join INFINITY LIFE and save lifes."
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
      TabIndex        =   5
      Top             =   1800
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   5655
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   14655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "New account"
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
      TabIndex        =   4
      Top             =   720
      Width           =   4095
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
      TabIndex        =   3
      Top             =   9600
      Width           =   20295
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
      Picture         =   "create_account.frx":6233
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
      Width           =   5655
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   255
      Left            =   9480
      Top             =   120
      Width           =   5295
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
Attribute VB_Name = "create_account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As String
Dim con As New ADODB.Connection
Private Sub Command1_Click()

If nametxt.Text = "" Or useridtxt.Text = "" Or passwordtxt.Text = "" Or addresstxt.Text = "" Or mobiletxt.Text = "" Then
MsgBox "Fill all the fields", vbCritical
ElseIf Check1.Value = 0 Then
MsgBox ("Confirm the entered details are true")
ElseIf Len(passwordtxt.Text) < 8 Then
Shape10.Visible = True
MsgBox "Weak password (minimum of 8 character)", vbExclamation
Else
Adodc1.RecordSource = "select * from user_account"
Adodc1.Recordset.AddNew
Adodc1.Recordset(0) = nametxt.Text
Adodc1.Recordset(1) = useridtxt.Text
Adodc1.Recordset(2) = passwordtxt.Text

If male.Value = True Then
Adodc1.Recordset(3) = male.Caption
ElseIf female.Value = True Then
Adodc1.Recordset(3) = female.Caption
ElseIf other.Value = True Then
Adodc1.Recordset(3) = other.Caption
End If

Adodc1.Recordset(4) = addresstxt.Text
Adodc1.Recordset(5) = emailtxt.Text
Adodc1.Recordset(6) = mobiletxt.Text

Adodc1.Recordset.Update
MsgBox "Your account is sucessfully created", vbInformation
nametxt.Text = ""
useridtxt.Text = ""
passwordtxt.Text = ""
addresstxt.Text = ""
emailtxt.Text = ""
mobiletxt.Text = ""
Shape10.Visible = False
male.Value = False
female.Value = False
other.Value = False
Check1.Value = 0
admin_home.Show
End If
End Sub

Private Sub Command2_Click()
user_details.Show
End Sub


Private Sub Form_Load()
cn = "Provider=MSDAORA.1;Password=open;User ID=scott;Persist Security Info=True"

End Sub

Private Sub gendertxt_keypress(Ascii As Integer)
If Ascii = 13 And gendertxt.Text <> "" Then
addresstxt.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub

Private Sub mobiletxt_keypress(Ascii As Integer)
If Ascii = 13 And mobiletxt.Text <> "" Then
mobiletxt.SetFocus
ElseIf (Ascii < 48 And Ascii <> 8) Or Ascii > 57 Then
Ascii = 0
MsgBox "Enter Numbers Only", vbInformation
End If
End Sub

Private Sub nametxt_keypress(Ascii As Integer)
If Ascii = 13 And nametxt.Text <> "" Then
useridtxt.SetFocus
ElseIf (Ascii < 65 And Ascii <> 8 And Ascii <> 32) Or (Ascii > 90 And Ascii < 97) Or (Ascii > 122) Then
Ascii = 0
MsgBox "Enter Letter Only", vbInformation
End If
End Sub
