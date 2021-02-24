VERSION 5.00
Begin VB.Form frmloginadmin 
   Appearance      =   0  'Flat
   BackColor       =   &H00404080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Login"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdadminlogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
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
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdadminback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtadminpass 
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
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtadminid 
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
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label lblpassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Admin ID"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lbladmintitle2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lbladmintitle1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ADMIN"
      BeginProperty Font 
         Name            =   "Segoe UI Black"
         Size            =   24
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmloginadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adminid As String
Dim adminpass As String
Dim adminidno As String

Private Sub cmdadminback_Click()
    frmindex.Show
    frmloginadmin.Hide
End Sub

Private Sub cmdadminlogin_Click()
    OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
    OpenAdminTable
    
    adminid = txtadminid
    rs3.Index = "admin_id"
    rs3.Seek "=", adminid

    If rs3.NoMatch Then
        MsgBox "Sorry no record found", vbOKOnly, "sorry"
    Else
        adminpass = rs3!admin_pass
        adminidno = rs3!admin_id
        If (adminidno = txtadminid And adminpass = txtadminpass) Then
            frmadminstudent.Show
            frmloginadmin.Hide
        Else
            MsgBox "Sorry wrong password", vbOKOnly, "sorry"
        End If
    End If
    txtadminid = ""
    txtadminpass = ""
End Sub


