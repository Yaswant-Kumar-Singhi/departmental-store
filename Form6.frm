VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form6"
   ClientHeight    =   7020
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   12855
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   21
      Text            =   "HERE"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Text            =   "WELCOME"
      Top             =   360
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10320
      Top             =   360
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0000FF00&
      Caption         =   "Login as Employee"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FF00&
      Caption         =   "Login As Admin"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Enter"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   3720
      TabIndex        =   24
      Top             =   1320
      Width           =   9255
   End
   Begin VB.Label Label18 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   3600
      TabIndex        =   23
      Top             =   6960
      Width           =   9255
   End
   Begin VB.Label Label17 
      BackColor       =   &H000080FF&
      Height          =   6975
      Left            =   12720
      TabIndex        =   22
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   9225
      Left            =   3720
      Picture         =   "Form6.frx":0000
      Top             =   1320
      Width           =   14595
   End
   Begin VB.Label Label15 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   19
      Top             =   5400
      Width           =   3615
   End
   Begin VB.Label Label14 
      Height          =   5775
      Left            =   3600
      TabIndex        =   18
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FF80&
      Caption         =   "LOGIN OPTION"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   15
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Height          =   2655
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiBold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FF80&
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      Height          =   2655
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label5 
      Height          =   1215
      Left            =   12480
      TabIndex        =   7
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label4 
      Height          =   1215
      Left            =   3600
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label3 
      Height          =   135
      Left            =   3600
      TabIndex        =   5
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label2 
      Height          =   135
      Left            =   3720
      TabIndex        =   4
      Top             =   1080
      Width           =   8775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "DEPARTMENTAL STORE SYSTEM"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label Label16 
      BackColor       =   &H0080FF80&
      Height          =   1575
      Left            =   0
      TabIndex        =   20
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Menu MENU 
      Caption         =   "MENU"
      Begin VB.Menu INFO 
         Caption         =   "INFO"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1 = True Then
Form1.Show
ElseIf Option2 = True Then
Form7.Show
Else
MsgBox ("Select as option")
End If






End Sub

Private Sub Command2_Click()
Form14.Show
End Sub

Private Sub Form_Load()
Label7.Caption = Format(Date, "dd-mmm-yyyy")
Label10.Caption = Format(Time, "hh:mm")
End Sub

Private Sub Info_Click()
Form15.Show

End Sub

Private Sub Timer1_Timer()
Label1.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label2.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label3.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label4.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label5.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label14.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))


End Sub
