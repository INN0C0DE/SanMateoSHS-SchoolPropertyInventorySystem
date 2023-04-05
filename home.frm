VERSION 5.00
Begin VB.Form home 
   BackColor       =   &H00FFC0C0&
   Caption         =   "San Mateo SHS - School Property Inventory System"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   Icon            =   "home.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   13365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "School Library"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "SCHOOL LIBRARY INVENTORY"
      Top             =   6120
      Width           =   3495
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "school clinic"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "SCHOOL CLINIC INVENTORY"
      Top             =   5160
      Width           =   3495
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "registrar's office"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "REGISTRAR'S OFFICE INVENTORY"
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SCHOOL CANTEEN"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "SCHOOL CANTEEN INVENTORY"
      Top             =   5160
      Width           =   3495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ACCOUNTING office"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "ACCOUNTING OFFICE INVENTORY"
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "principal's office"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "PRINCIPAL'S OFFICE INVENTORY"
      Top             =   3240
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "guidANCE office"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "GUIDANCE OFFICE INVENTORY"
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "faculty"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "FACULTY INVENTORY"
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "specialized rooms"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "SPECIALIZED ROOMS INVENTORY"
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "San Mateo SHS - School Property Inventory System"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12855
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Picture         =   "home.frx":2EEF2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Logout"
         Top             =   6600
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Students' room"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "STUDENTS' ROOM INVENTORY"
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Line Line8 
         X1              =   8160
         X2              =   9000
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line7 
         X1              =   4080
         X2              =   4800
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line4 
         X1              =   4080
         X2              =   4080
         Y1              =   1440
         Y2              =   6240
      End
      Begin VB.Line Line3 
         X1              =   9000
         X2              =   9000
         Y1              =   1440
         Y2              =   6240
      End
      Begin VB.Line Line2 
         X1              =   4080
         X2              =   4800
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         X1              =   8280
         X2              =   9000
         Y1              =   1440
         Y2              =   1440
      End
   End
   Begin VB.Line Line6 
      X1              =   6120
      X2              =   7320
      Y1              =   3600
      Y2              =   4080
   End
   Begin VB.Line Line5 
      X1              =   4320
      X2              =   4320
      Y1              =   3000
      Y2              =   3240
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BorderWidth     =   3
      Height          =   7695
      Left            =   0
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
MsgBox "Logging Off Complete!", vbInformation, "Status:"
loginfrm.Show
Me.Hide

End Sub

Private Sub Command10_Click()
regis.Show
Me.Hide
End Sub

Private Sub Command11_Click()
library.Show

End Sub

Private Sub Command2_Click()
tctfrm.Show
Me.Hide

End Sub

Private Sub Command3_Click()
labfrm.Show
Me.Hide
End Sub

Private Sub Command4_Click()
faculty.Show
Me.Hide

End Sub

Private Sub Command5_Click()
cpo.Show
Me.Hide

End Sub

Private Sub Command6_Click()
principal.Show
Me.Hide

End Sub

Private Sub Command7_Click()
acct.Show
Me.Hide
End Sub

Private Sub Command8_Click()
canteen.Show
Me.Hide

End Sub

Private Sub Command9_Click()
clinic.Show
Me.Hide

End Sub
