VERSION 5.00
Begin VB.Form frmsignupteacher 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   Caption         =   "Signup (Teacher)"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbackk 
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
      Left            =   6600
      TabIndex        =   27
      Top             =   8640
      Width           =   2895
   End
   Begin VB.CheckBox chktcrsc 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Science"
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
      Left            =   7200
      TabIndex        =   25
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CheckBox chktcrmat 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Mathematics"
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
      Left            =   7200
      TabIndex        =   24
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CheckBox chktcrbi 
      BackColor       =   &H00C0C0FF&
      Caption         =   "English"
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
      Left            =   5760
      TabIndex        =   23
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CheckBox chktcrbm 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Bahasa Malaysia"
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
      Left            =   5760
      TabIndex        =   22
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   7935
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   9135
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Subject"
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
         Height          =   2535
         Left            =   4680
         TabIndex        =   26
         Top             =   5160
         Width           =   3975
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
         TabIndex        =   11
         Top             =   720
         Width           =   8295
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
         TabIndex        =   10
         Top             =   1680
         Width           =   3135
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
         ItemData        =   "frmsignupteacher.frx":0000
         Left            =   3840
         List            =   "frmsignupteacher.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox cmbteacherstatus 
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
         ItemData        =   "frmsignupteacher.frx":001C
         Left            =   6240
         List            =   "frmsignupteacher.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   2055
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
         TabIndex        =   7
         Top             =   2640
         Width           =   8295
      End
      Begin VB.TextBox txtsignupteacheradress 
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
         Height          =   1215
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   3600
         Width           =   8295
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
         TabIndex        =   5
         Top             =   5280
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
         TabIndex        =   4
         Top             =   6240
         Width           =   3975
      End
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
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   7200
         Width           =   3975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label laberl 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   15
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
      Begin VB.Label Label9 
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
         Left            =   360
         TabIndex        =   13
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   360
         TabIndex        =   12
         Top             =   6840
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdsignuptaechersubmit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
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
      Left            =   600
      TabIndex        =   1
      Top             =   8640
      Width           =   2895
   End
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
      Left            =   3600
      TabIndex        =   0
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      TabIndex        =   21
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmsignupteacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resulttcr As Boolean
Dim ictcr As String
Dim idtcr As String

Private Sub cmdbackk_Click()
    frmloginteacher.Show
    Unload frmsignupteacher
End Sub

Private Sub cmdsignuptaechersubmit_Click()
    resulttcr = False
    Dim strtcr As String
    
    OpenTuitionDatabase ' call procedure openWasiatdatabase to open the Wasiat database
    OpenTeacherTable ' open student table in the Wasiat database file
    OpenTeachTempTable
    
    If (txtsignupteacherpassword = txtsignupteacherpasswordconfirm) Then
        rs2.AddNew
        rs2!tcr_name = UCase(txtsignupteacherfullname)
        rs2!tcr_ic = UCase(txtsignupteachericno)
        rs2!tcr_status = UCase(cmbteacherstatus)
        rs2!tcr_email = txtsignupteacheremail
        rs2!tcr_address = UCase(txtsignupteacheradress)
        rs2!tcr_phoneno = txtsignupteacberphoneno
        rs2!tcr_pass = txtsignupteacherpassword
        rs2!tcr_sex = cmbteachergender
        strtcr = "C:\Users\IQBAL\Desktop\PROJECT VB\defaultavatar.jpg"
        rs2!tcr_pic = strtcr
        rs2.Update
        'MsgBox "One record has been added ", vbOKOnly, "Add Record"
        
        ictcr = txtsignupteachericno
        rs2.Index = "tcr_ic"
        rs2.Seek "=", ictcr
        idtcr = rs2!tcr_id
        MsgBox "Your id is " & idtcr, vbOKCancel, "Add Record"
        
        If (chktcrbm.Value) Then
            rs9.AddNew
            subject = "BM01"
            rs9!sub_code = subject
            rs9!tcr_id = idtcr
            rs9.Update
        End If
        If (chktcrbi.Value) Then
            rs9.AddNew
            subject = "BI02"
            rs9!sub_code = subject
            rs9!tcr_id = idtcr
            rs9.Update
        End If
        If (chktcrmat.Value) Then
            rs9.AddNew
            subject = "MAT03"
            rs9!sub_code = subject
            rs9!tcr_id = idtcr
            rs9.Update
        End If
        If (chktcrsc.Value) Then
            rs9.AddNew
            subject = "SC04"
            rs9!sub_code = subject
            rs9!tcr_id = idtcr
            rs9.Update
        End If
        frmloginteacher.Show
        'frmsignupteacher.Hide
        Unload frmsignupteacher
    Else
        MsgBox "kata laluan tak sama", vbInformation, "try again"
        txtsignupteacherpassword = ""
        txtsignupteacherpasswordconfirm = ""
        txtsignupteacherpassword.SetFocus
    End If
'CloseEmpDatabase
End Sub

Private Sub cmdsignupteacherreset_Click()
    txtsignupteacherfullname = ""
    txtsignupteachericno = ""
    cmbteachergender.Clear
    cmbteacherstatus.Clear
    txtsignupteacheremail = ""
    txtsignupteacheradress = ""
    txtsignupteacberphoneno = ""
    txtsignupteacherpassword = ""
    txtsignupteacherpasswordconfirm = ""
    chktcrbm.Value = False
    chktcrmat.Value = False
    chktcrbi.Value = False
    chktcrsc.Value = False
    txtsignupteacherfullname.SetFocus
End Sub
