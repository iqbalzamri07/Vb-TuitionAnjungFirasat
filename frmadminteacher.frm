VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmadminteacher 
   Appearance      =   0  'Flat
   BackColor       =   &H00404080&
   BorderStyle     =   0  'None
   Caption         =   "Teacher Details (Admin)"
   ClientHeight    =   8970
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   7815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      ScaleHeight     =   795
      ScaleWidth      =   3915
      TabIndex        =   22
      Top             =   5400
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc adoteach 
      Height          =   330
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\IQBAL\Desktop\PROJECT VB\Tuition.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\IQBAL\Desktop\PROJECT VB\Tuition.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Teach"
      Caption         =   "teach"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoteacher 
      Height          =   495
      Left            =   5880
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\IQBAL\Desktop\PROJECT VB\Tuition.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\IQBAL\Desktop\PROJECT VB\Tuition.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Teacher"
      Caption         =   "teacher"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdcreatetimetabletcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CREATE TIME TABLE"
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
      Left            =   3000
      TabIndex        =   20
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtaddresstcr 
      Appearance      =   0  'Flat
      DataField       =   "tcr_address"
      DataSource      =   "adoteacher"
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
      Height          =   1095
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   6360
      Width           =   3975
   End
   Begin VB.CommandButton cmdadminteacherdelete 
      Appearance      =   0  'Flat
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   11.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   8160
      Width           =   3015
   End
   Begin VB.CommandButton cmdnexttcr 
      Appearance      =   0  'Flat
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Segoe UI Semilight"
         Size            =   11.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   17
      Top             =   8160
      Width           =   3135
   End
   Begin VB.TextBox txtidtcr 
      Appearance      =   0  'Flat
      DataField       =   "tcr_id"
      DataSource      =   "adoteacher"
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
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox txtictcr 
      Appearance      =   0  'Flat
      DataField       =   "tcr_ic"
      DataSource      =   "adoteacher"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox txtnametcr 
      Appearance      =   0  'Flat
      DataField       =   "tcr_name"
      DataSource      =   "adoteacher"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   3480
      Width           =   3975
   End
   Begin VB.TextBox txtstatustcr 
      Appearance      =   0  'Flat
      DataField       =   "tcr_status"
      DataSource      =   "adoteacher"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox txtemailtcr 
      Appearance      =   0  'Flat
      DataField       =   "tcr_email"
      DataSource      =   "adoteacher"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox txtsextcr 
      Appearance      =   0  'Flat
      DataField       =   "tcr_sex"
      DataSource      =   "adoteacher"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   4920
      Width           =   3975
   End
   Begin VB.TextBox txtphonenotcr 
      Appearance      =   0  'Flat
      DataField       =   "tcr_phoneno"
      DataSource      =   "adoteacher"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   7560
      Width           =   3975
   End
   Begin VB.TextBox txtsearchtcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton cmdsearchtcr 
      Appearance      =   0  'Flat
      Caption         =   "SEARCH ID"
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
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TEACHER DETAILS"
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
      Height          =   615
      Left            =   720
      TabIndex        =   23
      Top             =   840
      Width           =   6495
   End
   Begin VB.Label lblsubjectteacher 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   21
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lblidtcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblictcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   15
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblstatustcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label lblsextcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   13
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label lblnametcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblemailtcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lbladdresstcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label lblphonenotcr 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Menu mnuStudent 
      Caption         =   "STUDENT"
   End
   Begin VB.Menu mnuTeacher 
      Caption         =   "TEACHER"
   End
   Begin VB.Menu mnuTimeTable 
      Caption         =   "TIME TABLE"
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "LOG OUT"
   End
End
Attribute VB_Name = "frmadminteacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadminteacherdelete_Click()
    Dim strans, stricno As String

    OpenTuitionDatabase      'open the OpenWasiatDatabase function in the module
    OpenTeacherTable
    OpenTeachTable
    Dim cnn As New ADODB.Connection
    Dim rstdel As New ADODB.Recordset
    
    stridno = (txtidtcr)
    rs2.MoveFirst

    'nostaff adalah nama index key dalam staff table
    'to create index field , open table , click icon index , namakan filed index
    'selalunya field Primary Key

    rs2.Index = "tcr_id"
    rs2.Seek "=", stridno
    If rs2.NoMatch Then
     MsgBox "sorry no record found", vbOKOnly, "sorry"
    Else
    strans = MsgBox("Are Sure You Want To Delete This Record" & vbCrLf & "Name : " & (rs2.Fields("tcr_name").Value), vbYesNo, "Comfirmation")
    If strans = vbYes Then
        'rs2.Delete
        adoteacher.Recordset.Delete
        cmdnexttcr_Click
        adoteacher.Recordset.MoveFirst
        'rs7.MoveFirst
        Do
            rs7.Index = "tcr_id"
            rs7.Seek "=", stridno
            If rs7.NoMatch Then
                rs7.MoveNext
            Else
                rs7.Delete
            End If
        Loop Until rs7.EOF
        MsgBox "One Record Been Deleted", 16, "Record Delete"
    End If
        End If
        'clearteacher_click
        'frmadminstudent.Show
        rs2.MoveFirst
        rs2.Close
        Set rs2 = Nothing

End Sub

Private Sub cmdcreatetimetabletcr_Click()
    frmadmintimetabletcr.Show
    frmadminteacher.Hide
End Sub

Private Sub cmdnexttcr_Click()
    Dim tmp_id As String
    OpenTuitionDatabase
    OpenTeachTable
        adoteacher.Recordset.MoveNext
        If adoteacher.Recordset.EOF Then
            adoteacher.Recordset.MoveFirst
        End If
        
        Dim cnnn As New ADODB.Connection
        Dim rstsubtcrr As New ADODB.Recordset
        cnnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
        rstsubtcrr.Open "Select * FROM Teach;", cnnn, adOpenStatic
        tmp_id = txtidtcr
        rstsubtcrr.MoveFirst
        Pic.Cls
        Do
            If rstsubtcrr!tcr_id = tmp_id Then
                Pic.Print rstsubtcrr!sub_code
            End If
            rstsubtcrr.MoveNext
        Loop Until rstsubtcrr.EOF
    
End Sub

Private Sub cmdsearchtcr_Click()

    Dim tcridno As String
    
    OpenTuitionDatabase
    OpenTeacherTable
    
    tcridno = txtsearchtcr
    rs2.Index = "tcr_id"
    rs2.Seek "=", tcridno
    
    If rs2.NoMatch Then
        MsgBox "Sorry no record found", vbOKOnly, "sorry"
    Else
         txtidtcr = rs2!tcr_id
         txtictcr = rs2!tcr_ic
         txtnametcr = rs2!tcr_name
         txtstatustcr = rs2!tcr_status
         txtemailtcr = rs2!tcr_email
         txtsextcr = rs2!tcr_sex
         txtaddresstcr = rs2!tcr_address
         txtphonenotcr = rs2!tcr_phoneno
         pictcr = LoadPicture(rs2!tcr_pic)
         Dim tmp_id As String
            Dim cnnn As New ADODB.Connection
            Dim rstsubtcrr As New ADODB.Recordset
            cnnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
            rstsubtcrr.Open "Select * FROM Teach;", cnnn, adOpenStatic
            tmp_id = txtidtcr
            rstsubtcrr.MoveFirst
            Pic.Cls
            Do
                If rstsubtcrr!tcr_id = tmp_id Then
                    Pic.Print rstsubtcrr!sub_code
                End If
                rstsubtcrr.MoveNext
            Loop Until rstsubtcrr.EOF
    
End If

End Sub

Sub displaytcr()

Dim tcrid As String

OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
OpenTeacherTable
OpenTeachTable

     txtidtcr = rs2!tcr_id
     txtictcr = rs2!tcr_ic
     txtnametcr = rs2!tcr_name
     txtstatustcr = rs2!tcr_status
     txtemailtcr = rs2!tcr_email
     txtaddresstcr = rs2!tcr_address
     txtphonenotcr = rs2!tcr_phoneno
     pictcr = LoadPicture(rs2!tcr_pic)
     'txtsubjecttcr = rs2!tcr_subject

    tcrid = rs2!tcr_id
    Dim cnn As New ADODB.Connection
    Dim rstsubtcr As New ADODB.Recordset
    'Dim teacherID As String
    
    'teacherID = frmloginteacher.txttcrid
    
   
        
End Sub

Private Sub Form_Load()
    displaytcr
End Sub

Private Sub mnuLogout_Click()
    Unload frmadminteacher
    frmindex.Show
    
End Sub

Private Sub mnustudent_Click()
    frmadminstudent.Show
    'frmadminteacher.Hide
    Unload frmadminteacher
End Sub

Private Sub mnuteacher_Click()
    frmadminteacher.Show
    'frmadminstudent.Hide
    Unload frmadminstudent
End Sub

Private Sub mnuTimeTable_Click()
    frmadmintimetableteacher.Show
    'frmadminteacher.Hide
    Unload frmadminteacher
End Sub

Private Sub clearteacher_click()
    txtidtcr = ""
    txtictcr = ""
    txtnametcr = ""
    txtstatustcr = ""
    txtemailtcr = ""
    txtsextcr = ""
    listsubjectteacher = ""
    txtaddresstcr = ""
    txtphonenotcr = ""
End Sub
