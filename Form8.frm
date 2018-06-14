VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form8"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9780
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   8760
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Go to Login Portal"
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7800
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\pc\Desktop\New folder (2)\Database1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\pc\Desktop\New folder (2)\Database1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "T2"
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
      BackColor       =   &H0000FFFF&
      Caption         =   "BACK"
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "PASSWORD"
      DataSource      =   "Adodc1"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      DataField       =   "USERNAME"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6600
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "CONTACT NO"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2160
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "DATE OF BIRTH"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Register"
      Height          =   735
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "USERNAME AND PASSWORD ONCE CREATED CANNOT BE CHANGED.  "
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   2400
      Width           =   5775
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   135
      Left            =   5400
      TabIndex        =   14
      Top             =   2400
      Width           =   15
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT NO"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   150
      TabIndex        =   3
      Top             =   2580
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1890
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION DESK"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   11250
      Left            =   -2760
      Picture         =   "Form8.frx":0000
      Top             =   -1080
      Width           =   15000
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.Fields("NAME") = Text1.Text
Adodc1.Recordset.Fields("DATE OF BIRTH") = Text2.Text
Adodc1.Recordset.Fields("CONTACT NO") = Text3.Text
Adodc1.Recordset.Fields("USERNAME") = Text4.Text
Adodc1.Recordset.Fields("PASSWORD") = Text5.Text
Adodc1.Recordset.Update
MsgBox " User Registration Successfully done Login with Username and Password", vbInformation

End Sub

Private Sub Command2_Click()
Form7.Show
End Sub

Private Sub Command3_Click()
Form7.Show
Unload Me

End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
End Sub

Private Sub Timer1_Timer()
Dim str As String
str = Form8.Label8.Caption
str = Mid$(str, 2, Len(str)) + Left(str, 1)
Form8.Label8.Caption = str
End Sub
