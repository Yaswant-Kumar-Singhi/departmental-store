VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   " Login"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   7215
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         Caption         =   " Enter user name and Password Correctly to Login "
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   3570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   " Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "LOGIN"
      Height          =   255
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000A&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000FF&
      Height          =   2895
      Left            =   360
      Top             =   240
      Width           =   7935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FF00&
      Height          =   1575
      Index           =   4
      Left            =   480
      Top             =   1440
      Width           =   7695
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFF00&
      Height          =   495
      Index           =   3
      Left            =   1920
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFF00&
      FillColor       =   &H00C0C000&
      Height          =   495
      Index           =   2
      Left            =   1920
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000000FF&
      Height          =   1335
      Left            =   600
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000D&
      Height          =   495
      Index           =   1
      Left            =   6120
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H8000000D&
      Height          =   495
      Index           =   0
      Left            =   6120
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   1335
      Left            =   6120
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000D&
      Height          =   1095
      Left            =   600
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   975
      Left            =   480
      Top             =   360
      Width           =   7695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000011&
      Caption         =   " Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000011&
      Caption         =   " User Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   720
      Picture         =   "Form1.frx":0000
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1200
      Picture         =   "Form1.frx":0442
      Top             =   2040
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
rs.Close
rs.Source = "select * from T1 where Username = '" & Text1.Text & "' and Password = '" & Text2.Text & "'"
rs.Open
If (rs.EOF = True) Then
Text1.Text = ""
MsgBox " error"
ElseIf (rs.Fields(1).Value = Text2.Text) Then
Form2.Show
Unload Me
Else
MsgBox "invalid"
End If

End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Click Me. To move to ADMIN SECTION"
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Thanks,Now I am going to close this application"
End If
End Sub

Private Sub Form_Load()
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\pc\Desktop\New folder (2)\Database1.mdb;Persist Security Info=False"
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "T1"
rs.Open
Text1.Text = rs.Fields(1).Value

Label5.Caption = Format(Date, "dd-mmm-yyyy")




End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Enter the correct username to get control of admin section"
End If

End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Enter the correct password to get control of admin section"
End If
End Sub
