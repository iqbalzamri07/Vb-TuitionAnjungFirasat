VERSION 5.00
Begin VB.Form frmindex 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "index.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
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
      Left            =   6600
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdstudent 
      Appearance      =   0  'Flat
      Caption         =   "Student"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton cmdteacher 
      Appearance      =   0  'Flat
      Caption         =   "Teacher"
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
      Left            =   6600
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdadmin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Admin"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label lbltitle2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO "
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lbltitle1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PUSAT TUISYEN ANJUNG FIRASAT"
      BeginProperty Font 
         Name            =   "Segoe UI Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   4560
      TabIndex        =   0
      Top             =   960
      Width           =   6375
   End
End
Attribute VB_Name = "frmindex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadmin_Click()
    frmloginadmin.Show
    frmindex.Hide
End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdstudent_Click()
    frmloginstudent.Show
    frmindex.Hide
    
End Sub

Private Sub cmdteacher_Click()
    frmloginteacher.Show
    frmindex.Hide
End Sub
