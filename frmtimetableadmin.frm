VERSION 5.00
Begin VB.Form frmtimetableadmin 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Timetable (Admin)"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text28 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   27
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox Text27 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   26
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox Text26 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   25
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox Text25 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   24
      Top             =   7560
      Width           =   1695
   End
   Begin VB.TextBox Text24 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   23
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text23 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   22
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text22 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   21
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text21 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   20
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   19
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   18
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   17
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   16
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   15
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   14
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   13
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   12
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   9
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6720
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10.00 a.m - 11.00 a.m"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   4800
      TabIndex        =   44
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "08.00 p.m - 09.00 p.m"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   7080
      TabIndex        =   43
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "09.00 p.m - 10.00 p.m"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   9360
      TabIndex        =   42
      Top             =   1560
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
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   5760
      TabIndex        =   41
      Top             =   240
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
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   4560
      TabIndex        =   40
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblsun 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   39
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label lblsat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Saturday"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   38
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblfri 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Friday"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   37
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblthu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   36
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblwed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   35
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lbltue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   34
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblmon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Monday"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lbltime1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "09.00 a.m - 10.00 a.m"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   2520
      TabIndex        =   32
      Top             =   1560
      Width           =   1095
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   9360
      TabIndex        =   31
      Top             =   1200
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   7080
      TabIndex        =   30
      Top             =   1200
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4800
      TabIndex        =   29
      Top             =   1200
      Width           =   975
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2520
      TabIndex        =   28
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmtimetableadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
