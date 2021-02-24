VERSION 5.00
Begin VB.Form frmadminaddsubject 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Add Subject (Admin)"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   1890
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleMode       =   0  'User
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsubmit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   5
      Top             =   6720
      Width           =   2535
   End
   Begin VB.TextBox txtsubjectname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   4215
   End
   Begin VB.TextBox txtsubjectmodule 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   2400
      TabIndex        =   2
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox txtsubjectprice 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   2400
      TabIndex        =   1
      Top             =   5760
      Width           =   4215
   End
   Begin VB.TextBox txtsubjectcode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   2400
      TabIndex        =   0
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      Caption         =   "Subject Module"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      Caption         =   "Subject Price"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      Caption         =   "Subject Name"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      Caption         =   "Subject Code"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   8280
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lbladdsubject 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      Caption         =   "ADD SUBJECT"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000011&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Left            =   -600
      Shape           =   5  'Rounded Square
      Top             =   360
      Width           =   10215
   End
   Begin VB.Menu mnustudent 
      Caption         =   "STUDENT"
   End
   Begin VB.Menu mnuteacher 
      Caption         =   "TEACHER"
   End
   Begin VB.Menu mnutimetable 
      Caption         =   "TIME TABLE"
   End
   Begin VB.Menu mnuaddsubject 
      Caption         =   "ADD SUBJECT"
   End
End
Attribute VB_Name = "frmadminaddsubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuAddSubject_Click()
    frmadminaddsubject.Show
    frmadminstudent.Hide
    frmadminteacher.Hide
End Sub

Private Sub mnustudent_Click()
    frmadminaddsubject.Hide
    frmadminstudent.Show
    frmadminteacher.Hide
End Sub

Private Sub mnuteacher_Click()
    frmadminaddsubject.Hide
    frmadminteacher.Show
    frmadminstudent.Hide
End Sub

