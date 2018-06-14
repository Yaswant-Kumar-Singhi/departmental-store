VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form11"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6555
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Refresh"
      Height          =   435
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "QUIT"
      Height          =   435
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "LOGIN"
      Height          =   435
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Height          =   15
      Left            =   720
      TabIndex        =   14
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Height          =   15
      Left            =   840
      TabIndex        =   13
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Height          =   15
      Left            =   960
      TabIndex        =   12
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   840
      Width           =   15
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   5640
      TabIndex        =   8
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Height          =   1095
      Left            =   5400
      TabIndex        =   7
      Top             =   1560
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C000&
      Height          =   2055
      Left            =   720
      Top             =   720
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   1095
      Left            =   1200
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008080FF&
      Height          =   615
      Left            =   960
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "DEPARTMENTAL STORE MANAGEMENT"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = "store" And Text2.Text = "1" Then
Form6.Show
Unload Me
Else
MsgBox " Invalid Username or Password"
End If


End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Click Me. To open this application "
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Click Me. Toget out of this application"
End If
End Sub

Private Sub Command3_Click()
Form10.Show
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Well I am used to reload this application "
End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Enter correct username for opening this application "
End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Enter correct password for opening this application "
End If
End Sub
