VERSION 5.00
Begin VB.Form frmadmintimetableteacher 
   BackColor       =   &H00404080&
   Caption         =   "Time Table Teacher (Admin)"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   27
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   26
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   25
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   23
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   22
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   21
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      TabIndex        =   20
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   19
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   18
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   17
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      TabIndex        =   16
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   14
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   13
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      TabIndex        =   12
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   11
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   10
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   9
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      TabIndex        =   8
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   3
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox Text26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   2
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox Text27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6600
      TabIndex        =   1
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox Text28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      TabIndex        =   0
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TEACHER"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5880
      TabIndex        =   45
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblsession1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Session 1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblsession2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Session 2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4680
      TabIndex        =   43
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblsession3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Session 3"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6960
      TabIndex        =   42
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblsession4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Session 4"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9240
      TabIndex        =   41
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbltime1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "09.00 a.m - 10.00 a.m"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2400
      TabIndex        =   40
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblmon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   39
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbltue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   38
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label lblwed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   37
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblthu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   36
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblfri 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblsat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Saturday"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   34
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label lblsun 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   33
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3120
      TabIndex        =   32
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TABLE"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   24
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   4320
      TabIndex        =   31
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "09.00 p.m - 10.00 p.m"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   9240
      TabIndex        =   30
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "08.00 p.m - 09.00 p.m"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   6960
      TabIndex        =   29
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10.00 a.m - 11.00 a.m"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4680
      TabIndex        =   28
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Menu admintimetableteacher 
      Caption         =   "TEACHER"
   End
   Begin VB.Menu adminlogout 
      Caption         =   "LOG OUT"
   End
End
Attribute VB_Name = "frmadmintimetableteacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim teacherID As String

Private Sub adminlogout_Click()
    Unload frmadmintimetableteacher
    Unload frmadminteacher
    frmindex.Show
End Sub

Private Sub admintimetableteacher_Click()
    Unload frmadmintimetableteacher
    frmadminteacher.Show
            
        Dim tempp_id As String
        Dim cnnnnn As New ADODB.Connection
        Dim rstsubtcrrrr As New ADODB.Recordset
        
        cnnnnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
        rstsubtcrrrr.Open "Select * FROM Teach;", cnnnnn, adOpenStatic
        tempp_id = frmadminteacher.txtidtcr
        rstsubtcrrrr.MoveFirst
        frmadminteacher.Pic.Cls
        Do
            If rstsubtcrrrr!tcr_id = tempp_id Then
                frmadminteacher.Pic.Print rstsubtcrrrr!sub_code
            End If
            rstsubtcrrrr.MoveNext
        Loop Until rstsubtcrrrr.EOF
End Sub

Private Sub Form_Load()
    teacherID = frmadminteacher.txtidtcr
    OpenTuitionDatabase
    OpenTeachTable
    Dim cnn As New ADODB.Connection
    Dim rstbtcr As New ADODB.Recordset
    
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
    rstbtcr.Open "Select * FROM Teach;", cnn, adOpenStatic
    rstbtcr.MoveFirst
    Do
        If rstbtcr!tcr_id = teacherID Then
        
            If rstbtcr!timetable_day = "MONDAY" Then
                If rstbtcr!timetable_session = 1 Then
                    Text21 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 2 Then
                    Text22 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 3 Then
                    Text23 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 4 Then
                    Text24 = rstbtcr!sub_code
                End If
            ElseIf rstbtcr!timetable_day = "TUESDAY" Then
                If rstbtcr!timetable_session = 1 Then
                    Text1 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 2 Then
                    Text2 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 3 Then
                    Text3 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 4 Then
                    Text4 = rstbtcr!sub_code
                End If
            ElseIf rstbtcr!timetable_day = "WEDNESDAY" Then
                If rstbtcr!timetable_session = 1 Then
                    Text5 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 2 Then
                    Text6 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 3 Then
                    Text7 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 4 Then
                    Text8 = rstbtcr!sub_code
                End If
            ElseIf rstbtcr!timetable_day = "THURSDAY" Then
                If rstbtcr!timetable_session = 1 Then
                    Text9 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 2 Then
                    Text10 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 3 Then
                    Text11 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 4 Then
                    Text12 = rstbtcr!sub_code
                End If
            ElseIf rstbtcr!timetable_day = "FRIDAY" Then
                If rstbtcr!timetable_session = 1 Then
                    Text13 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 2 Then
                    Text14 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 3 Then
                    Text15 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 4 Then
                    Text16 = rstbtcr!sub_code
                End If
            ElseIf rstbtcr!timetable_day = "SATURDAY" Then
                If rstbtcr!timetable_session = 1 Then
                    Text17 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 2 Then
                    Text18 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 3 Then
                    Text19 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 4 Then
                    Text20 = rstbtcr!sub_code
                End If
            ElseIf rstbtcr!timetable_day = "SUNDAY" Then
                If rstbtcr!timetable_session = 1 Then
                    Text25 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 2 Then
                    Text26 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 3 Then
                    Text27 = rstbtcr!sub_code
                ElseIf rstbtcr!timetable_session = 4 Then
                    Text28 = rstbtcr!sub_code
                End If
            End If
        End If
        rstbtcr.MoveNext
    Loop Until rstbtcr.EOF
End Sub

