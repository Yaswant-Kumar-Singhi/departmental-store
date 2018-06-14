VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form7"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12135
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9360
      Top             =   5640
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   10920
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select *from T2"
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "CLICK ME"
      Height          =   255
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "LOGIN"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   8880
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   8880
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "FOR NEW REGISTRATION "
      Height          =   255
      Left            =   8640
      TabIndex        =   7
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label3 
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
      Left            =   7080
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "EMPLOYEE LOGIN PORTAL"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   6900
      Left            =   0
      Picture         =   "Form7.frx":0000
      Top             =   0
      Width           =   12285
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "Select* from T2 where Username='" + Text1.Text + "' and password='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Login Failed..Please Login with correct credentials", vbCritical
Form7.Show
Else
MsgBox "success"
Form9.Show
Form7.Hide
End If



End Sub

Private Sub Command2_Click()
Form6.Show
Unload Me
End Sub

Private Sub Command3_Click()
Form8.Show
Unload Me

End Sub

Private Sub Timer1_Timer()
Dim str As String
str = Form7.Label4.Caption
str = Mid$(str, 2, Len(str)) + Left(str, 1)
Form7.Label4.Caption = str
End Sub
