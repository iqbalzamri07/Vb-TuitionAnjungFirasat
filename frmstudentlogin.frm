VERSION 5.00
Begin VB.Form frmloginstudent 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Login"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsignupstd 
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
      Left            =   360
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdstdlogin 
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
      Left            =   6120
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdstdback 
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
      Left            =   4680
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtstdpass 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtstdid 
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
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblstdpass 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblstdid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Student ID"
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
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lbltitle2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lbltitle1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "STUDENT"
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
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "frmloginstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stdid As String
Dim stdpass As String
Dim stdidno As String

Private Sub cmdsignupstd_Click()
    frmsignupstudent.Show
    frmloginstudent.Hide
    
End Sub

Private Sub cmdstdback_Click()
    frmindex.Show
    frmloginstudent.Hide
End Sub

Private Sub cmdstdlogin_Click()
    OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
    OpenStudentTable
    
    stdid = txtstdid
    rs1.Index = "std_id"
    rs1.Seek "=", stdid

    If rs1.NoMatch Then
        MsgBox "Sorry no record found", vbOKOnly, "sorry"
    Else
        stdpass = rs1!std_pass
        stdidno = rs1!std_id
        If (stdidno = txtstdid And stdpass = txtstdpass) Then
            frmstudentdetails.Show
            'frmloginadmin.Hide
            Unload frmloginadmin
        Else
            MsgBox "Sorry wrong password", vbOKOnly, "sorry"
        End If
    End If
    txtstdid = ""
    txtstdpass = ""
    'CloseEmpDatabase
End Sub

