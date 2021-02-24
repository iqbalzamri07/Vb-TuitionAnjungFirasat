VERSION 5.00
Begin VB.Form frmloginteacher 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teacher Login"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsignuptcr 
      Appearance      =   0  'Flat
      Caption         =   "Sign Up"
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
      Left            =   480
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdtcrlogin 
      Appearance      =   0  'Flat
      Caption         =   "Login"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdtcrloginback 
      Appearance      =   0  'Flat
      Caption         =   "Back"
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
      Left            =   4560
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txttcrpass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txttcrid 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lbltcrpass 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lbltcrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Teacher ID"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lbltitle2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "LOG IN"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lbltitle1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "TEACHER"
      BeginProperty Font 
         Name            =   "Segoe UI Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmloginteacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim tcrpass As String
Dim tcridno As String

Private Sub cmdtcrback_Click()
    frmindex.Show
    frmloginteacher.Hide
End Sub

Private Sub cmdsignuptcr_Click()
    frmsignupteacher.Show
    'frmteacherlogin.Hide
End Sub

Private Sub cmdtcrlogin_Click()
    OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
    OpenTeacherTable
    Dim tcrid As String
    tcrid = ""
    tcrid = txttcrid
    rs2.Index = "tcr_id"
    rs2.Seek "=", tcrid

    If rs2.NoMatch Then
        MsgBox "Sorry no record found", vbOKOnly, "sorry"
    Else
        tcrpass = rs2!tcr_pass
        tcridno = rs2!tcr_id
        If (tcridno = txttcrid And tcrpass = txttcrpass) Then
            frmteacherdetails.Show
            Unload frmloginteacher
            'frmloginteacher.Hide
        Else
            MsgBox "Sorry wrong password", vbOKOnly, "sorry"
        End If
    End If
    txttcrid = ""
    txttcrpass = ""
    CloseEmpDatabase
End Sub

Private Sub cmdtcrloginback_Click()
    frmindex.Show
    frmloginteacher.Hide
End Sub
