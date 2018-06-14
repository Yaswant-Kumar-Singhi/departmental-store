VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H0080FF80&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ITEM NO"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8955
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "TRANS_DETAILS"
      Height          =   375
      Left            =   240
      TabIndex        =   45
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3120
      Top             =   600
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1920
      Top             =   480
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
      RecordSource    =   "TRANS"
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
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CALCULATE TOTAL"
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0000FF00&
      DataField       =   "TOTAL PRICE"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   7200
      TabIndex        =   35
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      Height          =   735
      Left            =   0
      TabIndex        =   28
      Top             =   6480
      Width           =   8895
      Begin VB.CommandButton Command13 
         BackColor       =   &H008080FF&
         Caption         =   "ITEM LIST"
         Height          =   375
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H008080FF&
         Caption         =   "STOCK WINDOW"
         Height          =   315
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H008080FF&
         Caption         =   "HR WINDOW"
         Height          =   315
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H008080FF&
         Caption         =   "MAIN WINDOW"
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H000000FF&
         Caption         =   "QUIT"
         Height          =   375
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "TIME"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFC0&
      DataField       =   "RATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFC0&
      DataField       =   "QUANTITY"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFC0&
      DataField       =   "ITEM NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFC0&
      DataField       =   "ITEM NO"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4320
      TabIndex        =   14
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      DataField       =   "TRASACTIONDATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080FFFF&
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "NEW"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "LAST"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5130
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "NEXT"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4500
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3750
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "FIRST"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   43
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "PURCHASE TRASACTION DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   37
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   5295
      Index           =   7
      Left            =   2280
      TabIndex        =   33
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   5295
      Index           =   6
      Left            =   8760
      TabIndex        =   32
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   135
      Index           =   5
      Left            =   2280
      TabIndex        =   31
      Top             =   6360
      Width           =   6615
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   135
      Index           =   4
      Left            =   2280
      TabIndex        =   30
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   5415
      Index           =   3
      Left            =   2040
      TabIndex        =   29
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   135
      Index           =   2
      Left            =   0
      TabIndex        =   27
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   26
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   25
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   8760
      TabIndex        =   23
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   1680
      TabIndex        =   22
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Height          =   135
      Left            =   1680
      TabIndex        =   21
      Top             =   720
      Width           =   7215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Height          =   135
      Left            =   1680
      TabIndex        =   20
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FF00&
      Height          =   1575
      Left            =   0
      TabIndex        =   19
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C000&
      Caption         =   "TOTAL PRICE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      Caption         =   "RATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "ITEM NAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "ITEM NO"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      Caption         =   "TRANSACTION DATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "STORE TRANSACTION RECORD"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FF00&
      Height          =   3735
      Left            =   0
      TabIndex        =   18
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label17 
      BackColor       =   &H0080FF80&
      Height          =   4935
      Left            =   2520
      TabIndex        =   34
      Top             =   1320
      Width           =   6135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Refresh

Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command10_Click()
Form2.Show
End Sub

Private Sub Command11_Click()
Form3.Show

End Sub

Private Sub Command12_Click()
Form5.Show
End Sub

Private Sub Command13_Click()
Form12.Show
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious

End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext

End Sub

Private Sub Command4_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveLast

End Sub

Private Sub Command5_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Command6_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""



End Sub

Private Sub Command8_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Text6.Text = Val(Text4.Text) * Val(Text5.Text)
End Sub

Private Sub Command9_Click()
Form16.Show

End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew

Label18.Caption = Format(Date, "dd-mm-yyyy")
Label19.Caption = Format(Time, "HH:MM")


End Sub

Private Sub Timer1_Timer()
Label11.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label12.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label13.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label14.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
End Sub
