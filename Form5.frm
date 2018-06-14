VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8565
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FF00&
      Caption         =   "OPEN TRANS WINDOW"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000FF00&
      Caption         =   "OPEN HR WINDOW"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "QUIT"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "LIST RECORD"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4680
      TabIndex        =   23
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Height          =   405
      Left            =   4680
      TabIndex        =   22
      Top             =   4920
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   1560
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
      CommandType     =   8
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
      RecordSource    =   ""
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
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   4680
      TabIndex        =   19
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   4680
      TabIndex        =   16
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   4680
      TabIndex        =   15
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "GO"
      Height          =   435
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   6240
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   480
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8040
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "UPDATE STOCK"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      TabIndex        =   40
      Top             =   6360
      Width           =   8415
      Begin VB.Label Label31 
         BackColor       =   &H0080FF80&
         Caption         =   "TIME"
         Height          =   255
         Left            =   6360
         TabIndex        =   44
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label30 
         BackColor       =   &H0080FF80&
         Caption         =   "DATE"
         Height          =   255
         Left            =   6360
         TabIndex        =   43
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label29 
         BackColor       =   &H0000FF00&
         Caption         =   "TIME"
         Height          =   255
         Left            =   7080
         TabIndex        =   42
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label28 
         BackColor       =   &H0000FF00&
         Caption         =   "DATE"
         Height          =   255
         Left            =   7080
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label27 
      BackColor       =   &H000080FF&
      Height          =   3975
      Left            =   8400
      TabIndex        =   39
      Top             =   2280
      Width           =   135
   End
   Begin VB.Label Label26 
      BackColor       =   &H000080FF&
      Height          =   4095
      Left            =   1920
      TabIndex        =   38
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label25 
      BackColor       =   &H000080FF&
      Height          =   3975
      Left            =   1680
      TabIndex        =   37
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label24 
      BackColor       =   &H000080FF&
      Height          =   3975
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label23 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   2040
      TabIndex        =   35
      Top             =   6120
      Width           =   6495
   End
   Begin VB.Label Label22 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   2040
      TabIndex        =   34
      Top             =   2160
      Width           =   6495
   End
   Begin VB.Label Label21 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   120
      TabIndex        =   33
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label20 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   120
      TabIndex        =   32
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label19 
      BackColor       =   &H000080FF&
      Height          =   975
      Left            =   8400
      TabIndex        =   31
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label18 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   120
      TabIndex        =   30
      Top             =   1920
      Width           =   8415
   End
   Begin VB.Label Label17 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   240
      TabIndex        =   29
      Top             =   1080
      Width           =   8295
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080FF80&
      Caption         =   "RACK NO"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   21
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      Caption         =   "UPDATE DATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "RATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FF80&
      Height          =   3975
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0FF&
      Height          =   855
      Left            =   8400
      TabIndex        =   9
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0FF&
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      Height          =   135
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      Height          =   135
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   8295
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      Caption         =   "ITEM NAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "ENTER THE ITEM NO TO UPDATE THE PRESENT STOCK"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "STOCK  UPDATION  PORTAL"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   8175
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      Height          =   3975
      Left            =   2040
      TabIndex        =   14
      Top             =   2280
      Width           =   6495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()
Adodc1.Recordset.Fields("ITEMNO") = Text1.Text
Adodc1.Recordset.Fields("ITEM NAME") = Text2.Text
Adodc1.Recordset.Fields("QUANTITY") = Text3.Text
Adodc1.Recordset.Fields("RATE") = Text4.Text
Adodc1.Recordset.Fields("STOCK") = Text5.Text
Adodc1.Recordset.Fields("UPDATE DATE") = Text6.Text
Adodc1.Recordset.Fields("RACK NO") = Text7.Text
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "Select * from HR where ITEMNO ='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Recod not Found,Please Try any other Item NO", vbInformation, "Message"
Else
Text2.Text = Adodc1.Recordset(1).Value
Text3.Text = Adodc1.Recordset(2).Value
Text4.Text = Adodc1.Recordset(3).Value
Text5.Text = Adodc1.Recordset(4).Value
Text6.Text = Adodc1.Recordset(5).Value
Text7.Text = Adodc1.Recordset(6).Value



End If
End Sub

Private Sub Command4_Click()
Form12.Show

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Form3.Show
    
End Sub

Private Sub Command7_Click()
Form4.Show

End Sub

Private Sub Form_Load()
Label28.Caption = Format(Date, "DD-MM-YYYY")
Label29.Caption = Format(Time, "HH:MM")
End Sub

Private Sub Timer1_Timer()
Label9.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label10.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label11.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label12.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))

End Sub

Private Sub Timer2_Timer()
Label1.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))

End Sub
