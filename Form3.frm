VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11550
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   10800
      Top             =   360
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5760
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Adodc2"
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
   Begin VB.TextBox Text8 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   5520
      TabIndex        =   24
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFC0&
      DataField       =   "RACK NO"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   8280
      TabIndex        =   23
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H0080FF80&
      DataField       =   "UPDATE DATE"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   8280
      TabIndex        =   22
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H0000FF00&
      DataField       =   "STOCK"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   5280
      TabIndex        =   21
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H0000FF00&
      DataField       =   "RATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0000FF00&
      DataField       =   "QUANTITY"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   5280
      TabIndex        =   19
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000FF00&
      DataField       =   "ITEM NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      DataField       =   "ITEMNO"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H000000FF&
      Caption         =   "QUIT"
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00808080&
      Caption         =   "SEARCH BY ITEM NO"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000080FF&
      Caption         =   "STOCK UPDATE"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF80&
      Caption         =   "DELETE"
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000080FF&
      Caption         =   "DO TRANSACTION"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "ADD RECORDDS"
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "LAST"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "NEXT"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "FIRST"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   5280
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
      RecordSource    =   "Select *from HR"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "MOVE TO "
      Height          =   1335
      Left            =   0
      TabIndex        =   33
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Caption         =   "NAVIGATOR"
      Height          =   4935
      Left            =   0
      TabIndex        =   34
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      Caption         =   "RECORD TAB"
      Height          =   1335
      Left            =   7440
      TabIndex        =   35
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label25 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   2400
      TabIndex        =   45
      Top             =   1080
      Width           =   9015
   End
   Begin VB.Label Label24 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   44
      Top             =   7560
      Width           =   11535
   End
   Begin VB.Label Label23 
      BackColor       =   &H000080FF&
      Height          =   7575
      Left            =   11400
      TabIndex        =   43
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label22 
      BackColor       =   &H000080FF&
      Height          =   1455
      Left            =   9840
      TabIndex        =   42
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label Label21 
      BackColor       =   &H000080FF&
      Height          =   1455
      Left            =   7320
      TabIndex        =   41
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label Label20 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   40
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   39
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackColor       =   &H000080FF&
      Height          =   7575
      Left            =   2400
      TabIndex        =   38
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   2520
      TabIndex        =   37
      Top             =   6120
      Width           =   9015
   End
   Begin VB.Label Label15 
      BackColor       =   &H0080FF80&
      Height          =   1335
      Left            =   9960
      TabIndex        =   36
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label17 
      BackColor       =   &H0080FF80&
      Height          =   1335
      Left            =   2520
      TabIndex        =   32
      Top             =   6240
      Width           =   4815
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FF80&
      Caption         =   "DATE"
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
      Left            =   240
      TabIndex        =   31
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FF80&
      Height          =   1095
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label12 
      Height          =   1095
      Left            =   11400
      TabIndex        =   29
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label11 
      Height          =   1095
      Left            =   2520
      TabIndex        =   28
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label10 
      Height          =   135
      Left            =   2640
      TabIndex        =   27
      Top             =   0
      Width           =   8775
   End
   Begin VB.Label Label9 
      Height          =   135
      Left            =   2640
      TabIndex        =   26
      Top             =   960
      Width           =   8775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "HANDLING RECORDS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   25
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RACK NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE DATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "RATE"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4530
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM NAME"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2790
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM NO"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   3660
      Width           =   2055
   End
   Begin VB.Label Label26 
      BackColor       =   &H0080FF80&
      Height          =   4695
      Left            =   7800
      TabIndex        =   46
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label27 
      BackColor       =   &H0000FF00&
      Caption         =   "ITEM DETAILS"
      Height          =   4695
      Left            =   2640
      TabIndex        =   47
      Top             =   1320
      Width           =   5055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc2.RecordSource = "Select * from HR where ITEMNO ='" + Text8.Text + "'"
Adodc2.Refresh
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command10_Click()
Adodc2.RecordSource = "Select * from HR where ITEMNO ='" + Text8.Text + "'"
Adodc2.Refresh
If Adodc2.Recordset.EOF Then
MsgBox "Recod not Found,Please Try any other Item NO", vbInformation, "Message"
Else
Text1.Text = Adodc2.Recordset(0).Value
Text2.Text = Adodc2.Recordset(1).Value
Text3.Text = Adodc2.Recordset(2).Value
Text4.Text = Adodc2.Recordset(3).Value
Text5.Text = Adodc2.Recordset(4).Value
Text6.Text = Adodc2.Recordset(5).Value
Text7.Text = Adodc2.Recordset(6).Value
End If

End Sub

Private Sub Command11_Click()
Unload Me

End Sub

Private Sub Command2_Click()
Adodc2.RecordSource = "Select * from HR where ITEMNO ='" + Text8.Text + "'"
Adodc2.Refresh
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command3_Click()
Adodc2.RecordSource = "Select * from HR where ITEMNO ='" + Text8.Text + "'"
Adodc2.Refresh
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub Command7_Click()
Form4.Show
End Sub

Private Sub Command8_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "Delete Record Confirmation")
If confirmation = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record Deleted Successfully", vbInformation, "Message"
Else
MsgBox "Record not deleted", vbInformation, "Message"
End If

End Sub

Private Sub Command9_Click()
Form5.Show

End Sub

Private Sub Image1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
Label14.Caption = Format(Date, " dd:mm:yyyy")
Label9.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label11.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
Label10.BackColor = (RGB(256 * Rnd, 256 * Rnd, 256 * Rnd))
End Sub

