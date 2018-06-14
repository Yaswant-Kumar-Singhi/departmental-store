VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form10"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4170
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   3000
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   390
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   1
      Max             =   500
   End
   Begin VB.Shape Shape7 
      Height          =   15
      Left            =   0
      Top             =   3960
      Width           =   4095
   End
   Begin VB.Shape Shape6 
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape5 
      Height          =   3975
      Left            =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Shape Shape4 
      Height          =   15
      Left            =   3480
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape3 
      Height          =   3495
      Left            =   3960
      Top             =   120
      Width           =   15
   End
   Begin VB.Shape Shape2 
      Height          =   15
      Left            =   3240
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   3720
      Top             =   120
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   " Departmental Store Project"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   3090
      Left            =   240
      Picture         =   "Form10.frx":0000
      Top             =   480
      Width           =   3660
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 10
    If ProgressBar1.Value = 500 Then
        Form11.Show
        Unload Me
        End If
        
End Sub
