VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form loginfrm 
   Caption         =   "San Mateo SHS - School Property Inventory System"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12195
   Icon            =   "loginfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adodata 
      Height          =   495
      Left            =   6240
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SMSHS - SPIS (Build 2.0)\Database\SMSHS - SPIS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SMSHS - SPIS (Build 2.0)\Database\SMSHS - SPIS.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "EXIT APPLICATION"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton loginbtn 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "LOGIN"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox password 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox username 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "© INN0C0DE 2020"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   10200
      TabIndex        =   8
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   555
      Left            =   120
      Picture         =   "loginfrm.frx":2EEF2
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   0
      Picture         =   "loginfrm.frx":31DF9
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   945
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4425
      Left            =   6480
      Picture         =   "loginfrm.frx":35277
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4545
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   11400
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "School Property Inventory System"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Top             =   5400
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "San Mateo Senior High School"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   4800
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   8415
      Left            =   5040
      Top             =   0
      Width           =   8775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   8415
      Left            =   0
      Top             =   -120
      Width           =   5055
   End
End
Attribute VB_Name = "loginfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cn As String
Dim bilang As Integer
Private Sub exitbtn_Click()
If MsgBox("Do you want to EXIT?", vbYesNo) = vbYes Then
    
    End
    End If
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
cn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SMSHS - SPIS (Build 2.0)\Database\SMSHS - SPIS.mdb;Persist Security Info=False"

End Sub

Private Sub loginbtn_Click()
If bilang > 2 Then
MsgBox "Please contact the System Administrator", vbCritical, "Invalid User:"
Unload Me
Exit Sub
End If

rs.Open "select * from SMSHS_Security where username='" & username.Text & "' and password='" & password.Text & "'", cn, 2, 2
If rs.EOF Then
    MsgBox "Access Denied!", vbInformation, "Please try again"
    username.Text = ""
    password.Text = ""
    username.SetFocus
    bilang = bilang + 1
    rs.Close
Else
    MsgBox "Access Granted!", vbInformation, "Login Status:"
    Unload Me
    smshs_loading.Show
    rs.Close
End If
End Sub
