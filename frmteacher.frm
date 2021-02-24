VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmteacherdetails 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Teacher Details"
   ClientHeight    =   8925
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox listsubjecttcr 
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
      Height          =   330
      ItemData        =   "frmteacher.frx":0000
      Left            =   2520
      List            =   "frmteacher.frx":0002
      TabIndex        =   17
      Top             =   5280
      Width           =   4335
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
      TabIndex        =   16
      Top             =   8160
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmduploadtcr 
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
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txttcrphoneno 
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
      Left            =   2520
      TabIndex        =   14
      Top             =   7560
      Width           =   4335
   End
   Begin VB.TextBox txttcraddress 
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
      Height          =   1650
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5760
      Width           =   4335
   End
   Begin VB.TextBox txttcremail 
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
      Left            =   2520
      TabIndex        =   12
      Top             =   4680
      Width           =   4335
   End
   Begin VB.TextBox txttcrstatus 
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
      Left            =   2520
      TabIndex        =   11
      Top             =   4080
      Width           =   4335
   End
   Begin VB.TextBox txttcrname 
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
      Left            =   2520
      TabIndex        =   10
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox txttcric 
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
      Left            =   2520
      TabIndex        =   9
      Top             =   2880
      Width           =   4335
   End
   Begin VB.TextBox txttcrid 
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
      Left            =   2520
      TabIndex        =   8
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Image pictcrr 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lbltcrphoneno 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label lbltcraddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label lbltcremail 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label lbltcrname 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lbltcrsubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "SUBJECT"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label lbltcrstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "STATUS"
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
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label lbltcric 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lbltcrid 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Menu mnuTcrDetail 
      Caption         =   "&TEACHER DETAILS"
      NegotiatePosition=   3  'Right
   End
   Begin VB.Menu mnuTimeTable 
      Caption         =   "TIME TA&BLE"
      NegotiatePosition=   3  'Right
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "LOG OUT"
   End
End
Attribute VB_Name = "frmteacherdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strtcr As String
Dim subjectCode As String

Private Sub cmdsearch_Click()
Dim idno As String

OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
OpenTeacherTable

idno = txttcrid
rs2.Index = "tcr_id"
rs2.Seek "=", idno
If rs1.NoMatch Then
   MsgBox "Sorry no record found", vbOKOnly, "sorry"
Else
   txttcrid = rs2!tcr_id
   txttcric = rs2!tcr_ic
   txttcrname = rs2!tcr_name
   txttcrstatus = rs2!tcr_status
   txttcremail = rs2!tcr_email
   txttcrsalary = rs2!tcr_salary
   txttcrsubject = rs2!tcr_subject
   txttcraddress = rs2!tcr_address
   txttcrhiredate = rs2!tcr_hiredate
   txttcrphoneno = rs2!tcr_phoneno
End If
'CloseEmpDatabase
End Sub

Private Sub cmdupdatestd_Click()
    OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
    OpenTeacherTable
    
    Dim strtcridno As String
    
        strtcridno = txttcrid
        rs2.Index = "tcr_id"
        rs2.Seek "=", strtcridno
        rs2.Edit
        rs2!tcr_ic = txttcric
        rs2!tcr_name = txttcrname
        rs2!tcr_status = txttcrstatus
        rs2!tcr_email = txttcremail
        'rs2!tcr_subject = txttcrsubject
        rs2!tcr_address = txttcraddress
        rs2!tcr_phoneno = txttcrphoneno
        rs2!tcr_pic = strtcr
        rs2.Update
        MsgBox "Record had been Updated ", vbOKOnly, "Add Record"
         'CloseEmpDatabase
End Sub

Private Sub cmduploadtcr_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Jpeg Files|*.jpg|GIF Files|*.*"
    CommonDialog1.ShowOpen
    strtcr = CommonDialog1.FileName
    pictcrr.Picture = LoadPicture(strtcr)
End Sub

Private Sub Form_Load()
    tcrdisplay
End Sub

Sub tcrdisplay()

    Dim tcrid As String
    
    OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
    OpenTeacherTable
    OpenTeachTable
    
     tcrid = frmloginteacher.txttcrid
     rs2.Index = "tcr_id"
     rs2.Seek "=", tcrid
     
     If rs2.NoMatch Then
        MsgBox "Sorry no record found", vbOKOnly, "sorry"
     Else
         txttcrid = rs2!tcr_id
         txttcric = rs2!tcr_ic
         txttcrname = rs2!tcr_name
         txttcrstatus = rs2!tcr_status
         txttcremail = rs2!tcr_email
         txttcraddress = rs2!tcr_address
         txttcrphoneno = rs2!tcr_phoneno
         pictcrr = LoadPicture(rs2!tcr_pic)
     End If
     
    Dim cnn As New ADODB.Connection
    Dim rstsubtcr As New ADODB.Recordset
    'Dim teacherID As String
    
    'teacherID = frmloginteacher.txttcrid
    
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
    rstsubtcr.Open "Select * FROM Teach;", cnn, adOpenStatic
    rstsubtcr.MoveFirst
        Do
            If rstsubtcr!tcr_id = tcrid Then
                listsubjecttcr.AddItem (rstsubtcr!sub_code)
            End If
            rstsubtcr.MoveNext
        Loop Until rstsubtcr.EOF
   'CloseEmpDatabase
End Sub

Private Sub mnuLogout_Click()
    'frmteacherdetails.Hide
    Unload frmteacherdetails
    frmindex.Show
    
End Sub


Private Sub mnuTimeTable_Click()
    frmteacherdetails.Hide
    frmtimetableteacher.Show
End Sub
