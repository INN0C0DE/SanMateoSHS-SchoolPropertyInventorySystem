VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form tctfrm 
   BackColor       =   &H00FFC0C0&
   Caption         =   " San Mateo SHS - School Property Inventory System"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   Icon            =   "tctfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
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
      Height          =   615
      Left            =   11280
      Picture         =   "tctfrm.frx":2EEF2
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "PRINT DATA"
      Top             =   6240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Students' Room"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
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
         Height          =   495
         Left            =   12720
         Picture         =   "tctfrm.frx":326F4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Back"
         Top             =   7920
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         DataField       =   "Track:"
         DataSource      =   "adodata"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   20
         Text            =   "SELECT TRACK"
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "No of Chairs:"
         DataSource      =   "adodata"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "No of Tables:"
         DataSource      =   "adodata"
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
         Left            =   2520
         TabIndex        =   18
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "No of TV:"
         DataSource      =   "adodata"
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
         Left            =   2520
         TabIndex        =   17
         Top             =   4800
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         DataField       =   "Strand:"
         DataSource      =   "adodata"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   16
         Text            =   "SELECT STRAND"
         Top             =   2520
         Width           =   4215
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         DataField       =   "Section:"
         DataSource      =   "adodata"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   15
         Text            =   "SELECT SECTION"
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         DataField       =   "Room_No:"
         DataSource      =   "adodata"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         Left            =   9480
         Picture         =   "tctfrm.frx":328EF
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "SAVE/UPDATE"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
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
         Height          =   610
         Left            =   8640
         Picture         =   "tctfrm.frx":360F1
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "ADD NEW"
         Top             =   6120
         Width           =   735
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
         Left            =   10320
         Picture         =   "tctfrm.frx":398F3
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "DELETE"
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         DataField       =   "No of E-fan:"
         DataSource      =   "adodata"
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
         Left            =   2520
         TabIndex        =   10
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataField       =   "No of Blackboard:"
         DataSource      =   "adodata"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   6000
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         Appearance      =   0  'Flat
         DataField       =   "Grade Level:"
         DataSource      =   "adodata"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   8
         Text            =   "SELECT GRADE LEVEL"
         Top             =   1560
         Width           =   4215
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
         Left            =   10560
         Picture         =   "tctfrm.frx":3D0F5
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5280
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
         Left            =   9360
         Picture         =   "tctfrm.frx":3D2E8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5280
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         DataField       =   "Others:"
         DataSource      =   "adodata"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2520
         TabIndex        =   4
         Top             =   6600
         Width           =   4335
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   1680
         Picture         =   "tctfrm.frx":3D4D6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "HELP"
         Top             =   6720
         Width           =   375
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
         Left            =   9600
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton searchbtn 
         Appearance      =   0  'Flat
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
         Left            =   10920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tctfrm.frx":40CD8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Search"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin MSAdodcLib.Adodc adodata 
         Height          =   495
         Left            =   4560
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
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
         RecordSource    =   "Students_Room"
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
         Bindings        =   "tctfrm.frx":40EA0
         Height          =   3615
         Left            =   7200
         TabIndex        =   7
         Top             =   1560
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
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
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "Room_No:"
            Caption         =   "Room No:"
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
            DataField       =   "Grade Level:"
            Caption         =   "Grade Level:"
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
            DataField       =   "Track:"
            Caption         =   "Track:"
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
            DataField       =   "Strand:"
            Caption         =   "Strand:"
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
            DataField       =   "Section:"
            Caption         =   "Section:"
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
            DataField       =   "No of Chairs:"
            Caption         =   "No of Chairs:"
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
         BeginProperty Column06 
            DataField       =   "No of Tables:"
            Caption         =   "No of Tables:"
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
         BeginProperty Column07 
            DataField       =   "No of TV:"
            Caption         =   "No of TV:"
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
         BeginProperty Column08 
            DataField       =   "No of E-fan:"
            Caption         =   "No of E-fan:"
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
         BeginProperty Column09 
            DataField       =   "No of Blackboard:"
            Caption         =   "No of Blackboard:"
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
         BeginProperty Column10 
            DataField       =   "Others:"
            Caption         =   "Others:"
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
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Track:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   33
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Chairs:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   32
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Tables:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   31
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of TV:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   30
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Strand:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   29
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Section:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   28
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Room No:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   27
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of E-fan:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   26
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Blackboard:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   25
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade Level:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Others:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Find Room No. :"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
   End
End
Attribute VB_Name = "tctfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Combo2.Clear
If Combo1.Text = "ACADEMIC" Then
Combo2.AddItem "STEM"
Combo2.AddItem "HUMMS"
Combo2.AddItem "ABM"
Combo2.AddItem "GAS"
ElseIf Combo1.Text = "TECH-VOC" Then
Combo2.AddItem "COMPUTER PROGRAMMING"
Combo2.AddItem "ANIMATION"
Combo2.AddItem "EIM"
Combo2.AddItem "COOKERY/FBS"
Combo2.AddItem "BEAUTY CARE"
Else
End If
End Sub

Private Sub Combo2_Click()
Combo3.Clear
If Combo2.Text = "STEM" Then
Combo3.AddItem "STEM"
ElseIf Combo2.Text = "HUMMS" Then
Combo3.AddItem "HUMMS-A"
Combo3.AddItem "HUMMS-B"
Combo3.AddItem "HUMMS-C"
ElseIf Combo2.Text = "ABM" Then
Combo3.AddItem "ABM-A"
Combo3.AddItem "ABM-B"
ElseIf Combo2.Text = "GAS" Then
Combo3.AddItem "GAS"
ElseIf Combo2.Text = "COMPUTER PROGRAMMING" Then
Combo3.AddItem "COMPUTER PROGRAMMING"
ElseIf Combo2.Text = "ANIMATION" Then
Combo3.AddItem "ANIMATION"
ElseIf Combo2.Text = "EIM" Then
Combo3.AddItem "EIM"
ElseIf Combo2.Text = "COOKERY/FBS" Then
Combo3.AddItem "COOKERY/FBS-A"
Combo3.AddItem "COOKERY/FBS-B"
ElseIf Combo2.Text = "BEAUTY CARE" Then
Combo3.AddItem "BEAUTY CARE"
End If
End Sub



Private Sub Combo4_Click()
Combo1.Clear
If Combo4.Text = "GRADE 11" Then
Combo1.AddItem "ACADEMIC"
Combo1.AddItem "TECH-VOC"
ElseIf Combo4.Text = "GRADE 12" Then
Combo1.AddItem "ACADEMIC"
Combo1.AddItem "TECH-VOC"
Else
End If
End Sub

Private Sub Command1_Click()
home.Show
Me.Hide

End Sub

Private Sub Command2_Click()
MsgBox "Data has been saved successfully!", vbInformation, "Saving Status:"
adodata.Recordset.Update
End Sub

Private Sub Command3_Click()
adodata.Recordset.AddNew
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

Private Sub Command7_Click()
MsgBox "Put/Type the name of the item as well as the number of item/s it has", vbInformation, "Help:"
End Sub

Private Sub Command8_Click()
students_report.Show


End Sub

Private Sub Form_Load()


Combo4.AddItem "GRADE 11"
Combo4.AddItem "GRADE 12"
End Sub

Private Sub searchbtn_Click()
If search.Text = "" Then
MsgBox "Error", vbCritical, "Something is missing..."
adodata.Refresh
Else
adodata.Refresh
adodata.Recordset.Find "Room_No: ='" & search.Text & "'"

End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Else
KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
Else
KeyAscii = 0
End If
End Sub



