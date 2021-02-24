VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsignupteacherreset 
      Appearance      =   0  'Flat
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   21
      Top             =   8280
      Width           =   3975
   End
   Begin VB.CommandButton cmdsignuptaechersubmit 
      Appearance      =   0  'Flat
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   20
      Top             =   8280
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Teacher Sign Up"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   8895
      Begin VB.TextBox txtsignupteacherpasswordconfirm 
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
         Height          =   495
         Left            =   4680
         TabIndex        =   19
         Top             =   6360
         Width           =   3975
      End
      Begin VB.TextBox txtsignupteacherpassword 
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
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   6360
         Width           =   3975
      End
      Begin VB.TextBox txtsignupteacberphoneno 
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
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   5280
         Width           =   3975
      End
      Begin VB.TextBox txtsignupteacheradress 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   3600
         Width           =   8295
      End
      Begin VB.TextBox txtsignupteacheremail 
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
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   8295
      End
      Begin VB.ComboBox cmbteacherstatus 
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
         ItemData        =   "teacherSignup.frx":0000
         Left            =   6240
         List            =   "teacherSignup.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox cmbteachergender 
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
         ItemData        =   "teacherSignup.frx":0035
         Left            =   3840
         List            =   "teacherSignup.frx":003F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtsignupteachericno 
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
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtsignupteacherfullname 
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
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   8295
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4680
         TabIndex        =   17
         Top             =   6000
         Width           =   2415
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Marital Status"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label laberl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "IC Number"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   11.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PUSAT TUISYEN ANJUNG FIRASAT"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
