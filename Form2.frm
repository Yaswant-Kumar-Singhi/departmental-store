VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   5370
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9165
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Stock_Update"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Transaction"
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Handling_Records"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   5700
      Left            =   -6360
      Picture         =   "Form2.frx":0000
      Top             =   -120
      Width           =   17445
   End
   Begin VB.Menu MENU 
      Caption         =   "MENU"
      Begin VB.Menu Handling_Records 
         Caption         =   "Handling Records"
         Shortcut        =   {F1}
      End
      Begin VB.Menu STR_TRANS_RCD 
         Caption         =   "Store Transaction Record"
         Shortcut        =   {F2}
      End
      Begin VB.Menu StockUpdate 
         Caption         =   "Stock  Update"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu Info 
         Caption         =   "Info"
         Shortcut        =   ^I
      End
      Begin VB.Menu Calculator 
         Caption         =   "Calculator"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub


Private Sub Calculator_Click()
Form13.Show

End Sub

Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub Command2_Click()
Form4.Show
End Sub

Private Sub Command3_Click()
Form5.Show
End Sub

Private Sub Command5_Click()
Form6.Show
Unload Me
End Sub

Private Sub Exit_Click()
End

End Sub

Private Sub Handling_Records_Click()
Form3.Show
End Sub

Private Sub Info_Click()
Form15.Show

End Sub

Private Sub SimpleCalculator_Click()

End Sub

Private Sub StockUpdate_Click()
Form5.Show

End Sub

Private Sub STR_TRANS_RCD_Click()
Form4.Show
End Sub
