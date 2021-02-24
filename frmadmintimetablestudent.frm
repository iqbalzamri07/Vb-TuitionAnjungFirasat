VERSION 5.00
Begin VB.Form frmadmintimetablestudent 
   BackColor       =   &H00404080&
   Caption         =   "Time Table Student (Admin)"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   11280
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
      Left            =   2160
      TabIndex        =   27
      Top             =   3360
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
      Left            =   4440
      TabIndex        =   26
      Top             =   3360
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
      Left            =   6720
      TabIndex        =   25
      Top             =   3360
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
      Left            =   9000
      TabIndex        =   24
      Top             =   3360
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
      Left            =   2160
      TabIndex        =   23
      Top             =   4200
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
      Left            =   4440
      TabIndex        =   22
      Top             =   4200
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
      Left            =   6720
      TabIndex        =   21
      Top             =   4200
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
      Left            =   9000
      TabIndex        =   20
      Top             =   4200
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
      Left            =   2160
      TabIndex        =   19
      Top             =   5040
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
      Left            =   4440
      TabIndex        =   18
      Top             =   5040
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
      Left            =   6720
      TabIndex        =   17
      Top             =   5040
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
      Left            =   9000
      TabIndex        =   16
      Top             =   5040
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
      Left            =   2160
      TabIndex        =   15
      Top             =   5880
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
      Left            =   4440
      TabIndex        =   14
      Top             =   5880
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
      Left            =   6720
      TabIndex        =   13
      Top             =   5880
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
      Left            =   9000
      TabIndex        =   12
      Top             =   5880
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
      Left            =   2160
      TabIndex        =   11
      Top             =   6720
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
      Left            =   4440
      TabIndex        =   10
      Top             =   6720
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
      Left            =   6720
      TabIndex        =   9
      Top             =   6720
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
      Left            =   9000
      TabIndex        =   8
      Top             =   6720
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
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
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
      Left            =   4440
      TabIndex        =   6
      Top             =   2520
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
      Left            =   6720
      TabIndex        =   5
      Top             =   2520
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
      Left            =   9000
      TabIndex        =   4
      Top             =   2520
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
      Left            =   2160
      TabIndex        =   3
      Top             =   7560
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
      Left            =   4440
      TabIndex        =   2
      Top             =   7560
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
      Left            =   6720
      TabIndex        =   1
      Top             =   7560
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
      Left            =   9000
      TabIndex        =   0
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT"
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
      Width           =   2295
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
      TabIndex        =   44
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
      TabIndex        =   43
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
      TabIndex        =   42
      Top             =   1200
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   9360
      TabIndex        =   41
      Top             =   1200
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
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   2520
      TabIndex        =   40
      Top             =   1560
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   39
      Top             =   2520
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   38
      Top             =   3360
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   37
      Top             =   4200
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   36
      Top             =   5040
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   35
      Top             =   5880
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   34
      Top             =   6720
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
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   7560
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
      Left            =   3240
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
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   4440
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
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   9360
      TabIndex        =   30
      Top             =   1560
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
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   7080
      TabIndex        =   29
      Top             =   1560
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
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   4800
      TabIndex        =   28
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Menu adminstudent 
      Caption         =   "STUDENT"
   End
   Begin VB.Menu adminlogoutstd 
      Caption         =   "LOG OUT"
   End
End
Attribute VB_Name = "frmadmintimetablestudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim studentID As String

Private Sub adminlogoutstd_Click()
    Unload frmadminstudent
    Unload frmadmintimetablestudent
    frmindex.Show
End Sub

Private Sub adminstudent_Click()
    Unload frmadmintimetablestudent
    frmadminstudent.Show
End Sub

Private Sub Form_Load()
    studentID = frmadminstudent.txtidstd
    OpenTuitionDatabase
    OpenLearnTable
    Dim cnn As New ADODB.Connection
    Dim rstbstd As New ADODB.Recordset
    
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
    rstbstd.Open "Select * FROM Learn;", cnn, adOpenStatic
    rstbstd.MoveFirst
    Do
        If rstbstd!std_id = studentID Then
        
            If rstbstd!timetable_day = "MONDAY" Then
                If rstbstd!timetable_session = 1 Then
                    Text21 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 2 Then
                    Text22 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 3 Then
                    Text23 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 4 Then
                    Text24 = rstbstd!sub_code
                End If
            ElseIf rstbstd!timetable_day = "TUESDAY" Then
                If rstbstd!timetable_session = 1 Then
                    Text1 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 2 Then
                    Text2 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 3 Then
                    Text3 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 4 Then
                    Text4 = rstbstd!sub_code
                End If
            ElseIf rstbstd!timetable_day = "WEDNESDAY" Then
                If rstbstd!timetable_session = 1 Then
                    Text5 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 2 Then
                    Text6 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 3 Then
                    Text7 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 4 Then
                    Text8 = rstbstd!sub_code
                End If
            ElseIf rstbstd!timetable_day = "THURSDAY" Then
                If rstbstd!timetable_session = 1 Then
                    Text9 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 2 Then
                    Text10 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 3 Then
                    Text11 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 4 Then
                    Text12 = rstbstd!sub_code
                End If
            ElseIf rstbstd!timetable_day = "FRIDAY" Then
                If rstbstd!timetable_session = 1 Then
                    Text13 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 2 Then
                    Text14 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 3 Then
                    Text15 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 4 Then
                    Text16 = rstbstd!sub_code
                End If
            ElseIf rstbstd!timetable_day = "SATURDAY" Then
                If rstbstd!timetable_session = 1 Then
                    Text17 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 2 Then
                    Text18 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 3 Then
                    Text19 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 4 Then
                    Text20 = rstbstd!sub_code
                End If
            ElseIf rstbstd!timetable_day = "SUNDAY" Then
                If rstbstd!timetable_session = 1 Then
                    Text25 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 2 Then
                    Text26 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 3 Then
                    Text27 = rstbstd!sub_code
                ElseIf rstbstd!timetable_session = 4 Then
                    Text28 = rstbstd!sub_code
                End If
            End If
        End If
        rstbstd.MoveNext
    Loop Until rstbstd.EOF
End Sub

