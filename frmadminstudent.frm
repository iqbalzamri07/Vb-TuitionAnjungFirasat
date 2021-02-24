VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmadminstudent 
   Appearance      =   0  'Flat
   BackColor       =   &H00404080&
   BorderStyle     =   0  'None
   Caption         =   "Student Details (Admin)"
   ClientHeight    =   8565
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc adostudent 
      Height          =   615
      Left            =   5400
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
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
      RecordSource    =   "Student"
      Caption         =   "adostudent"
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
   Begin VB.CommandButton cmdcreatetimetablestd 
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
      Left            =   3120
      TabIndex        =   22
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtaddressstd 
      Appearance      =   0  'Flat
      DataField       =   "std_address"
      DataSource      =   "adostudent"
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
      Height          =   1215
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   5280
      Width           =   3975
   End
   Begin VB.CommandButton cmdadminstddelete 
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
      TabIndex        =   20
      Top             =   7800
      Width           =   3015
   End
   Begin VB.CommandButton cmdnextstd 
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
      Left            =   720
      TabIndex        =   19
      Top             =   7800
      Width           =   3135
   End
   Begin VB.TextBox txtidstd 
      Appearance      =   0  'Flat
      DataField       =   "std_id"
      DataSource      =   "adostudent"
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
      TabIndex        =   9
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox txticstd 
      Appearance      =   0  'Flat
      DataField       =   "std_ic"
      DataSource      =   "adostudent"
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
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox txtnamestd 
      Appearance      =   0  'Flat
      DataField       =   "std_name"
      DataSource      =   "adostudent"
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
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txtsexstd 
      Appearance      =   0  'Flat
      DataField       =   "std_sex"
      DataSource      =   "adostudent"
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
      Top             =   3840
      Width           =   3975
   End
   Begin VB.TextBox txtemailstd 
      Appearance      =   0  'Flat
      DataField       =   "std_email"
      DataSource      =   "adostudent"
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
      Top             =   4320
      Width           =   3975
   End
   Begin VB.TextBox txtcategorystd 
      Appearance      =   0  'Flat
      DataField       =   "std_category"
      DataSource      =   "adostudent"
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
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox txtparentnostd 
      Appearance      =   0  'Flat
      DataField       =   "std_parentno"
      DataSource      =   "adostudent"
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
      Top             =   6600
      Width           =   3975
   End
   Begin VB.TextBox txtphonenostd 
      Appearance      =   0  'Flat
      DataField       =   "std_phoneno"
      DataSource      =   "adostudent"
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
      Top             =   7080
      Width           =   3975
   End
   Begin VB.CommandButton cmdsearchstd 
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
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtsearchstd 
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT DETAILS"
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
   Begin VB.Label lblidstd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   18
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblicstd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   17
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblsexstd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   16
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label lbladdressstd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   1
      Left            =   720
      TabIndex        =   15
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label lblcategorystd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label lblnamestd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   13
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblemailstd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   12
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblparentnostd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   11
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label lblphonenostd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   10
      Top             =   7080
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
Attribute VB_Name = "frmadminstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadminstddelete_Click()
    Dim strstd, strstdicno As String

    OpenTuitionDatabase      'open the OpenWasiatDatabase function in the module
    OpenStudentTable
    OpenLearnTable
    
    strstdicno = (txtidstd)
    rs1.MoveFirst

    'nostaff adalah nama index key dalam staff table
    'to create index field , open table , click icon index , namakan filed index
    'selalunya field Primary Key

    rs1.Index = "std_id"
    rs1.Seek "=", strstdicno
    If rs1.NoMatch Then
     MsgBox "sorry no record found", vbOKOnly, "sorry"
    Else
    strstd = MsgBox("Are Sure You Want To Delete This Record" & vbCrLf & "Name : " & (rs1.Fields("std_name").Value), vbYesNo, "Comfirmation")
    If strstd = vbYes Then
        adostudent.Recordset.Delete
        cmdnextstd_Click
        adostudent.Recordset.MoveFirst
        'rs5.MoveFirst
        Do
            rs5.Index = "std_id"
            rs5.Seek "=", strstdicno
            If rs5.NoMatch Then
                rs5.MoveNext
            Else
                rs5.Delete
                'adostudent.Recordset.Delete
            End If
        Loop Until rs5.EOF
        MsgBox "One Record Been Deleted", 16, "Record Delete"
    End If
        End If
        rs1.Close
        Set rs1 = Nothing
End Sub

Private Sub cmdcreatetimetablestd_Click()
    frmadmintimetablestd.Show
    frmadminstudent.Hide
    'Unload frmadminstudent
End Sub

Private Sub cmdnextstd_Click()

    adostudent.Recordset.MoveNext
    If adostudent.Recordset.EOF Then
        adostudent.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdsearchstd_Click()
Dim stdidno As String

OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
OpenStudentTable

stdidno = txtsearchstd
rs1.Index = "std_id"
rs1.Seek "=", stdidno

If rs1.NoMatch Then
    MsgBox "Sorry no record found", vbOKOnly, "sorry"
Else
    txtidstd = rs1!std_id
    txticstd = rs1!std_ic
    txtnamestd = rs1!std_name
    txtsexstd = rs1!std_sex
    txtemailstd = rs1!std_email
    txtcategorystd = rs1!std_category
    txtaddressstd = rs1!std_address
    txtphonenostd = rs1!std_phoneno
    picstd = LoadPicture(rs1!std_pic)
End If

End Sub

Sub displaystd()

OpenTuitionDatabase   'open the OpenWasiatDatabase function in the module
OpenStudentTable

    txtidstd = rs1!std_id
    txticstd = rs1!std_ic
    txtnamestd = rs1!std_name
    cmbgender = rs1!std_sex
    txtemailstd = rs1!std_email
    cmbcategory = rs1!std_category
    txtaddressstd = rs1!std_address
    txtparentnostd = rs1!std_parentno
    txtphonenostd = rs1!std_phoneno
    picstd = LoadPicture(rs1!std_pic)

End Sub

Private Sub Form_Load()
    displaystd
End Sub

Private Sub mnuLogout_Click()
    Unload frmadminstudent
    frmindex.Show
    
End Sub

Private Sub mnustudent_Click()
    'frmadminaddsubject.Hide
    frmadminstudent.Show
    'frmadminteacher.Hide
    Unload frmadminteacher
End Sub

Private Sub mnuteacher_Click()
    frmadminteacher.Show
    'frmadminstudent.Hide
    Unload frmadminstudent
        Dim temp_id As String
        Dim cnnnn As New ADODB.Connection
        Dim rstsubtcrrr As New ADODB.Recordset
        cnnnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
        rstsubtcrrr.Open "Select * FROM Teach;", cnnnn, adOpenStatic
        temp_id = frmadminteacher.txtidtcr
        rstsubtcrrr.MoveFirst
        frmadminteacher.Pic.Cls
        Do
            If rstsubtcrrr!tcr_id = temp_id Then
                frmadminteacher.Pic.Print rstsubtcrrr!sub_code
            End If
            rstsubtcrrr.MoveNext
        Loop Until rstsubtcrrr.EOF
End Sub

Private Sub mnuTimeTable_Click()
    frmadmintimetablestudent.Show
    'frmadminstudent.Hide
    Unload frmadminstudent
    
End Sub

Private Sub clearstudent_click()
    txtidstd = ""
    txticstd = ""
    txtnamestd = ""
    txtsexstd = ""
    txtemailstd = ""
    txtcategorystd = ""
    txtaddressstd = ""
    txtparentnostd = ""
    txtphonenostd = ""
End Sub
