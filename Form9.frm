VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form9"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9690
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "DLT_PRICE"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   3720
      TabIndex        =   50
      Top             =   5280
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "QUIT"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6720
      TabIndex        =   39
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Use Calulator"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CLEAR "
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "DELETE ITEM"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ADD LIST"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3000
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   2280
      TabIndex        =   34
      Top             =   2760
      Width           =   5655
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000C000&
      Caption         =   "cal Price"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1560
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8160
      Top             =   2160
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
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "PRINT"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "VIEW  ITEM"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "NEW BILL"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   7080
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   5640
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "CLICK"
      Height          =   435
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   0
      TabIndex        =   45
      Top             =   6600
      Width           =   9615
      Begin VB.Label Label31 
         BackColor       =   &H0080FF80&
         Caption         =   "TIME"
         Height          =   255
         Left            =   8160
         TabIndex        =   48
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label32 
         BackColor       =   &H0080FF80&
         Caption         =   "TIME"
         Height          =   255
         Left            =   8640
         TabIndex        =   47
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label30 
         BackColor       =   &H0000FF00&
         Caption         =   "Billing System Ver_01.02.13.28."
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label Label29 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   8040
      TabIndex        =   44
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label28 
      BackColor       =   &H000080FF&
      Height          =   3015
      Left            =   8040
      TabIndex        =   43
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label27 
      BackColor       =   &H000080FF&
      Height          =   3015
      Left            =   9480
      TabIndex        =   42
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label26 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   8040
      TabIndex        =   41
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000FF00&
      Height          =   3135
      Left            =   8040
      TabIndex        =   40
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label24 
      BackColor       =   &H0000FF00&
      Caption         =   "TOTAL "
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5520
      TabIndex        =   29
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label23 
      BackColor       =   &H000080FF&
      Height          =   1455
      Left            =   9480
      TabIndex        =   28
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label22 
      BackColor       =   &H000080FF&
      Height          =   1455
      Left            =   2280
      TabIndex        =   27
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label21 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   2280
      TabIndex        =   26
      Top             =   2520
      Width           =   7335
   End
   Begin VB.Label Label20 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   2280
      TabIndex        =   25
      Top             =   1200
      Width           =   7335
   End
   Begin VB.Label Label19 
      BackColor       =   &H000080FF&
      Height          =   5175
      Left            =   0
      TabIndex        =   24
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label18 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   23
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H000080FF&
      Height          =   5175
      Left            =   2040
      TabIndex        =   22
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Label16 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   21
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackColor       =   &H000080FF&
      Height          =   135
      Left            =   0
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000FF00&
      Height          =   3135
      Left            =   0
      TabIndex        =   19
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000FF00&
      Caption         =   "Price"
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
      Left            =   7080
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      Height          =   1095
      Left            =   9600
      TabIndex        =   14
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0FF&
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Height          =   135
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9735
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   9735
   End
   Begin VB.Label Label6 
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "BILLING SYSTEM"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FF00&
      Caption         =   "Quantity"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Rate"
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
      Left            =   5640
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Item Name"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Item No"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000FF00&
      Caption         =   "        ITEM FINDER"
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FF80&
      Height          =   1335
      Left            =   2280
      TabIndex        =   18
      Top             =   1320
      Width           =   7335
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Adodc1.RecordSource = "Select * from HR where ITEMNO ='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Recod not Found,Please Try any other Item NO", vbInformation, "Message"
Else
Text2.Text = Adodc1.Recordset(1).Value
Text3.Text = Adodc1.Recordset(2).Value
Text4.Text = Adodc1.Recordset(3).Value


End If

End Sub



Private Sub Command10_Click()
List1.Clear
Text6.Text = ""
Text7.Text = ""
End Sub



Private Sub Command11_Click()
Text6.Text = Val(Text6.Text) - Val(Text7.Text)

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

List1.Clear

End Sub

Private Sub Command3_Click()
Form13.Show
End Sub


Private Sub Command4_Click()
Form12.Show
End Sub

Private Sub Command5_Click()
CommonDialog1.ShowPrinter
PrintForm





End Sub

Private Sub Command6_Click()
Text5.Text = Val(Text3.Text) * Val(Text4.Text)

If Command6 = True Then
Text6.Text = Val(Text6.Text) + Val(Text5.Text)
End If


End Sub

Private Sub Command7_Click()
Form6.Show
Unload Me
End Sub

Private Sub Command8_Click()
 List1.AddItem ("ITEM NO : " + Text1.Text + "   " + "ITEM NAME :" + "   " + Text2.Text + "    " + "QUANTITY : " + "   " + Text3.Text + "  " + "RATE: " + "   " + Text4.Text + "  " + "PRICE : " + "   " + Text5.Text)
 
End Sub


Private Sub Command9_Click()
Dim ind As Integer
ind = List1.ListIndex
If ind >= 0 Then
List1.RemoveItem (ind)
End If




End Sub

Private Sub Form_Load()
Label6.Caption = Format(Date, "dd-mm-yyyy")

Label32.Caption = Format(Time, "hh:mm")



End Sub



Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
MsgBox "Here You can the items purchased by the customer"
End If
End Sub

Private Sub Text5_Change()


a = Val(Text3.Text * Text4.Text)
Text5.Text = Val(a)
End Sub
