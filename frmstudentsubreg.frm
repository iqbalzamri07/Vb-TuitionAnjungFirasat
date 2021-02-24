VERSION 5.00
Begin VB.Form frmstudentsubreg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Subject Registration"
   ClientHeight    =   6135
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Segoe UI Light"
      Size            =   11.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6135
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frasubreg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Subject Registration"
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3855
      Begin VB.CheckBox chkbm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bahasa Melayu"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkbi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "English"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox chksc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Science"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox chkmt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mathematics"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         TabIndex        =   2
         Top             =   3240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdsubjectsubmit 
      Appearance      =   0  'Flat
      Caption         =   "Submit"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Menu stddetails 
      Caption         =   "STUDENT DETAI&LS"
   End
   Begin VB.Menu stdsubreg 
      Caption         =   "SUB&JECT REGISTRATION"
   End
   Begin VB.Menu stdtimetable 
      Caption         =   "TIME&TABLE"
   End
End
Attribute VB_Name = "frmstudentsubreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
