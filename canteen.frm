VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form canteen 
   BackColor       =   &H00FFC0C0&
   Caption         =   "San Mateo SHS - School Property Inventory System"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   Icon            =   "canteen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "SMSHS Canteen :"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   13095
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
         Left            =   3360
         Picture         =   "canteen.frx":2EEF2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "PRINT DATA"
         Top             =   6600
         Width           =   855
      End
      Begin VB.CommandButton searchbtn 
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
         Height          =   495
         Left            =   10440
         Picture         =   "canteen.frx":326F4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6600
         Width           =   615
      End
      Begin VB.TextBox search 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   7
         Top             =   6600
         Width           =   1935
      End
      Begin VB.CommandButton Command8 
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
         Height          =   495
         Left            =   12360
         Picture         =   "canteen.frx":328BC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Back"
         Top             =   7080
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000000&
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
         Left            =   1440
         Picture         =   "canteen.frx":32AB7
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "ADD ITEM"
         Top             =   6600
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H80000000&
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
         Left            =   2400
         Picture         =   "canteen.frx":362B9
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "DELETE"
         Top             =   6600
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000000&
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
         Left            =   6960
         Picture         =   "canteen.frx":39ABB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6600
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H80000000&
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
         Left            =   5640
         Picture         =   "canteen.frx":39CAE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6600
         Width           =   615
      End
      Begin MSAdodcLib.Adodc adodata 
         Height          =   375
         Left            =   5400
         Top             =   7200
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SMSHS - SPIS (Build 2.0)\Database\SMSHS - SPIS.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\SMSHS - SPIS (Build 2.0)\Database\SMSHS - SPIS.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SMSHS_Canteen"
         Caption         =   "Database"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "canteen.frx":39E9C
         Height          =   6015
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   10610
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   14737632
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   24
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Room No :"
            Caption         =   "Room No :"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Item_Name:"
            Caption         =   "Item Name:"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "No of Item/s :"
            Caption         =   "No of Item/s :"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Date Recorded :"
            Caption         =   "Date Recorded :"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Name of Custodian :"
            Caption         =   "Name of Custodian :"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Remarks:"
            Caption         =   "Remarks:"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2085.166
            EndProperty
         EndProperty
      End
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   8175
      Left            =   0
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "canteen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
canteen_report.Show

End Sub

Private Sub Command3_Click()
canteen1.Show

End Sub

Private Sub Command4_Click()
adodata.Recordset.Delete
End Sub

Private Sub Command5_Click()
On Error Resume Next
adodata.Recordset.MoveNext

End Sub

Private Sub Command6_Click()
On Error Resume Next
adodata.Recordset.MovePrevious
End Sub

Private Sub Command8_Click()
home.Show
Me.Hide
End Sub

Private Sub searchbtn_Click()
If search.Text = "" Then
MsgBox "Error", vbCritical, "Something is missing..."
adodata.Refresh
Else
adodata.Refresh
adodata.Recordset.Find "Item_Name: ='" & search.Text & "'"

End If
End Sub
