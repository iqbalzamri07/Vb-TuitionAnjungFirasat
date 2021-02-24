VERSION 5.00
Begin VB.Form frmsignupstudent 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Signup (Student)"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdback 
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
      TabIndex        =   29
      Top             =   8880
      Width           =   2775
   End
   Begin VB.CommandButton cmdsignupreset 
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
      TabIndex        =   23
      Top             =   8880
      Width           =   2895
   End
   Begin VB.CommandButton cmdsignupsubmit 
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
      Left            =   720
      TabIndex        =   22
      Top             =   8880
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Student Sign Up"
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
      Height          =   8175
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   8895
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   1815
         Left            =   4680
         TabIndex        =   24
         Top             =   6120
         Width           =   3735
         Begin VB.CheckBox chkstdsc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1800
            TabIndex        =   28
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CheckBox chkstdbi 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   480
            TabIndex        =   27
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox chkstdmat 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   1800
            TabIndex        =   26
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkstdbm 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   480
            TabIndex        =   25
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.TextBox txtsignuppasswordconfirm 
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
         TabIndex        =   21
         Top             =   7440
         Width           =   3975
      End
      Begin VB.TextBox txtsignuppassword 
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
         TabIndex        =   20
         Top             =   6360
         Width           =   3975
      End
      Begin VB.TextBox txtsignupparentphoneno 
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
         TabIndex        =   17
         Top             =   5280
         Width           =   3975
      End
      Begin VB.TextBox txtsignupphoneno 
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
      Begin VB.TextBox txtsignupadress 
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
         TabIndex        =   13
         Top             =   3600
         Width           =   8295
      End
      Begin VB.TextBox txtsignupemail 
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
      Begin VB.ComboBox cmbcategory 
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
         ItemData        =   "frmsignupstd.frx":0000
         Left            =   6240
         List            =   "frmsignupstd.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox cmbgender 
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
         ItemData        =   "frmsignupstd.frx":0062
         Left            =   3840
         List            =   "frmsignupstd.frx":006C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtsignupicno 
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
      Begin VB.TextBox txtsignupfullname 
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
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   19
         Top             =   7080
         Width           =   2415
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   18
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Parent Phone Number"
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
         TabIndex        =   16
         Top             =   4920
         Width           =   2655
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         BackColor       =   &H00FFFFFF&
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
         BackColor       =   &H00FFFFFF&
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
         BackColor       =   &H00FFFFFF&
         Caption         =   "Category"
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
         BackColor       =   &H00FFFFFF&
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
         BackColor       =   &H00FFFFFF&
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
         BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmsignupstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resultstd As Boolean
Dim icstd As String
Dim idstd As String
Dim subject As String
Dim countsub As Integer

Private Sub cmdback_Click()
    frmloginstudent.Show
    Unload frmsignupstudent
End Sub

Private Sub cmdsignupreset_Click()
    txtsignupfullname = ""
    txtsignupicno = ""
    cmbgender.Clear
    cmbcategory.Clear
    txtsignupemail = ""
    txtsignupadress = ""
    txtsignupparentphoneno = ""
    txtsignupphoneno = ""
    txtsignuppassword = ""
    txtsignuppasswordconfirm = ""
    chkstdbm.Value = False
    chkstdmat.Value = False
    chkstdbi.Value = False
    chkstdsc.Value = False
    txtsignupfullname.SetFocus
End Sub

Private Sub cmdsignupsubmit_Click()
    resultstd = False
    Dim strdtd As String
    countsub = 0
    OpenTuitionDatabase ' call procedure openWasiatdatabase to open the Wasiat database
    OpenStudentTable ' open student table in the Wasiat database file
    OpenLearnTempTable
    
     If (txtsignuppassword = txtsignuppasswordconfirm) Then
         rs1.AddNew
         rs1!std_name = UCase(txtsignupfullname)
         rs1!std_ic = UCase(txtsignupicno)
         rs1!std_category = cmbcategory
         rs1!std_email = txtsignupemail
         rs1!std_address = UCase(txtsignupadress)
         rs1!std_phoneno = txtsignupphoneno
         rs1!std_pass = txtsignuppassword
         rs1!std_sex = cmbgender
         rs1!std_parentno = txtsignupparentphoneno
         'strstd = "C:\Users\IQBAL\Desktop\PROJECT VB\defaultavatar.jpg"
         rs1!std_pic = "C:\Users\IQBAL\Desktop\PROJECT VB\defaultavatar.jpg"
         rs1.Update
         'MsgBox "Your id ", vbOKOnly, "Add Record"
         
         icstd = txtsignupicno
         rs1.Index = "std_ic"
         rs1.Seek "=", icstd
         idstd = rs1!std_id
         MsgBox "Your id is " & idstd, vbOKOnly, "Add Record"
         
         Open "C:/Users/IQBAL/Desktop/PROJECT VB/receipt.txt" For Output As #1
         Write #1, "=========PUSAT TUISYEN ANJUNG FIRASAT========="
         Write #1, "==================RECEIPT====================="
         Write #1, "             Subject Taken:                            "
         If (chkstdbm.Value) Then
             rs8.AddNew
             subject = "BM01"
             rs8!sub_code = subject
             rs8!std_id = idstd
             rs8.Update
             Write #1, "             BM01    : RM20                    "
             countsub = countsub + 20
         End If
         If (chkstdbi.Value) Then
             rs8.AddNew
             subject = "BI02"
             rs8!sub_code = subject
             rs8!std_id = idstd
             rs8.Update
             Write #1, "             BI02    : RM20                    "
             countsub = countsub + 20
         End If
         If (chkstdmat.Value) Then
             rs8.AddNew
             subject = "MAT03"
             rs8!sub_code = subject
             rs8!std_id = idstd
             rs8.Update
             Write #1, "             MAT03   : RM20                   "
             countsub = countsub + 20
         End If
         If (chkstdsc.Value) Then
             rs8.AddNew
             subject = "SC04"
             rs8!sub_code = subject
             rs8!std_id = idstd
             rs8.Update
             Write #1, "             SC04    : RM20                    "
             countsub = countsub + 20
         End If
         Dim printtotal As String
         printtotal = "             TOTAL   : RM"
         Write #1, printtotal; countsub
         Close #1
         frmloginstudent.Show
         'frmsignupstudent.Hide
         Unload frmsignupstudent
     Else
         MsgBox "kata laluan tak sama", vbInformation, "try again"
         txtsignuppassword = ""
         txtsignuppasswordconfirm = ""
         txtsignuppassword.SetFocus
     End If
     
End Sub

