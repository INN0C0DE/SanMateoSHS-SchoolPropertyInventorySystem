VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form smshs_loading 
   BackColor       =   &H80000002&
   Caption         =   "San Mateo SHS - School Property Inventory System"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   5400
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2040
      Top             =   5400
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
      Left            =   10080
      TabIndex        =   3
      Top             =   6120
      Width           =   1815
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
      Left            =   3480
      TabIndex        =   1
      Top             =   4200
      Width           =   5415
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
      Left            =   3960
      TabIndex        =   0
      Top             =   4800
      Width           =   4575
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   9000
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4515
      Left            =   3960
      Picture         =   "smshs_loading.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4740
   End
End
Attribute VB_Name = "smshs_loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If ProgressBar1.Value <> ProgressBar1.Max Then
    ProgressBar1.Value = ProgressBar1.Value + 1
Else
    Unload Me
    home.Show
End If
End Sub
