VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmstudentdetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Student Details"
   ClientHeight    =   9285
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmduploadstd 
      Appearance      =   0  'Flat
      Caption         =   "UPLOAD"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdupdatestd 
      Appearance      =   0  'Flat
      Caption         =   "Update"
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
      Left            =   2520
      TabIndex        =   18
      Top             =   8640
      Width           =   2415
   End
   Begin VB.TextBox txtstdphoneparentdetails 
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
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   8160
      Width           =   4215
   End
   Begin VB.TextBox txtstdphonedetails 
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
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   7560
      Width           =   4215
   End
   Begin VB.TextBox txtstdemaildetails 
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
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   3960
      Width           =   4215
   End
   Begin VB.TextBox txtstdcategorydetails 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   5160
      Width           =   4215
   End
   Begin VB.TextBox txtstdaddressdetails 
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
      Height          =   1575
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5760
      Width           =   4215
   End
   Begin VB.TextBox txtstdgenderdetails 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   4560
      Width           =   4215
   End
   Begin VB.TextBox txtstdnamedetails 
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
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   3360
      Width           =   4215
   End
   Begin VB.TextBox txtstdicdetails 
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
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   2760
      Width           =   4215
   End
   Begin VB.TextBox txtstdiddetails 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Image picstd 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblparentphonenumstd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PARENT NUMBER"
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
      Left            =   720
      TabIndex        =   8
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label lblphonenumstd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PHONE NUMBER"
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
      Left            =   720
      TabIndex        =   7
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label lblemailstd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "EMAIL"
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
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblcategorystd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CATEGORY"
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
      Left            =   720
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label lbladdressstd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ADDRESS"
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
      Left            =   720
      TabIndex        =   4
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label lblgenderstd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GENDER"
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
      Left            =   720
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblnamestd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "NAME"
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
      Left            =   720
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblicstd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "IC"
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
      Left            =   720
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label lblidstd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ID"
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
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Menu stddetails 
      Caption         =   "STUDENT DETAI&LS"
   End
   Begin VB.Menu stdtimetable 
      Caption         =   "TIME&TABLE"
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "LOG OUT"
   End
End
Attribute VB_Name = "frmstudentdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strstd As String

Sub displaystd()

    OpenTuitionDatabase   'open the OpenStudentDatabase function in the module
    OpenStudentTable
    
    stdid = frmloginstudent.txtstdid
    rs1.Index = "std_id"
    rs1.Seek "=", stdid
    
    If rs1.NoMatch Then
        MsgBox "Sorry no record found", vbOKOnly, "sorry"
    Else
        txtstdiddetails = rs1!std_id
        txtstdicdetails = rs1!std_ic
        txtstdnamedetails = rs1!std_name
        txtstdgenderdetails = rs1!std_sex
        txtstdemaildetails = rs1!std_email
        txtstdcategorydetails = rs1!std_category
        txtstdaddressdetails = rs1!std_address
        txtstdphonedetails = rs1!std_phoneno
        txtstdphoneparentdetails = rs1!std_parentno
        picstd = LoadPicture(rs1!std_pic)
    End If
    'CloseEmpDatabase
End Sub

Private Sub cmdupdatestd_Click()
    OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
    OpenStudentTable
    
    Dim strstdidno As String
    
        strstdidno = txtstdiddetails
        rs1.Index = "std_id"
        rs1.Seek "=", strstdidno
        rs1.Edit
        rs1!std_ic = txtstdicdetails
        rs1!std_name = txtstdnamedetails
        rs1!std_phoneno = txtstdphonedetails
        rs1!std_email = txtstdemaildetails
        'rs1!std_subject = txttcrsubject
        rs1!std_address = txtstdaddressdetails
        rs1!std_parentno = txtstdphoneparentdetails
        rs1!std_pic = strstd
        rs1.Update
        MsgBox "Record had been Updated ", vbOKOnly, "Add Record"
        
        'CloseEmpDatabase
End Sub

Private Sub cmduploadstd_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Jpeg Files|*.jpg|GIF Files|*.*"
    CommonDialog1.ShowOpen
    strstd = CommonDialog1.FileName
    picstd.Picture = LoadPicture(strstd)
End Sub

Private Sub Form_Load()
    displaystd
    frmloginstudent.Hide
End Sub

Private Sub mnuLogout_Click()
    'frmstudentdetails.Hide
    Unload frmstudentdetails
    frmindex.Show
    
End Sub

Private Sub stdtimetable_Click()
        frmtimetablestd.Show
        frmstudentdetails.Hide
End Sub
