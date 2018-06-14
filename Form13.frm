VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      Height          =   405
      Left            =   225
      TabIndex        =   7
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   3735
      Left            =   120
      Top             =   840
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As String
Dim op1 As Double, op2 As Double, res As Double
Private Sub Command1_Click(Index As Integer)
Text1.Text = Text1.Text + Command1(Index).Caption
End Sub

Private Sub Command2_Click(Index As Integer)
op1 = Val(Text1.Text)
Text1.Text = Clear
r = Command2(Index).Caption
End Sub

Private Sub Command3_Click()
op2 = Val(Text1.Text)
Text1.Text = Clear
If (r = "+") Then
Text1.Text = op1 + op2
ElseIf (r = "-") Then
Text1.Text = op1 - op2
ElseIf (r = "*") Then
Text1.Text = op1 * op2
Else
Text1.Text = op1 / op2
End If
End Sub

Private Sub Command4_Click()
 Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
End Sub

Private Sub Command5_Click()
Text1.Text = Clear
End Sub




