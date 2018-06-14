VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   LinkTopic       =   "Form14"
   ScaleHeight     =   3645
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   1800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Saving..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   3090
      Left            =   0
      Picture         =   "Form14.frx":0000
      Top             =   0
      Width           =   3660
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 10
    If ProgressBar1.Value = 500 Then
        End
        End If
End Sub
