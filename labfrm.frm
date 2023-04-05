VERSION 5.00
Begin VB.Form labfrm 
   BackColor       =   &H00FFC0C0&
   Caption         =   "San Mateo SHS - School Property Inventory System"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   Icon            =   "labfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "San Mateo SHS - Specialized Rooms :"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   12855
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "electrical shop"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Electrical Shop Inventory"
         Top             =   4080
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "computer laboratory"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Computer Laboratory Inventory"
         Top             =   2040
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "food labORATORY"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Food Laboratory Inventory"
         Top             =   4080
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Science laboratory"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Science Laboratory Inventory"
         Top             =   2040
         UseMaskColor    =   -1  'True
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
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
      Height          =   615
      Left            =   480
      Picture         =   "labfrm.frx":2EEF2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Home"
      Top             =   7200
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   8175
      Left            =   0
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "labfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
home.Show
Me.Hide

End Sub

Private Sub Command2_Click()
scilab.Show
Me.Hide

End Sub

Private Sub Command3_Click()
foodlab.Show
Me.Hide

End Sub

Private Sub Command4_Click()
comlab.Show
Me.Hide

End Sub

Private Sub Command5_Click()
elecshop.Show
Me.Hide

End Sub
