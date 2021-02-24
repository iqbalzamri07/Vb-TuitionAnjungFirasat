VERSION 5.00
Begin VB.Form frmadmintimetablestd 
   BackColor       =   &H00404080&
   Caption         =   "Student Subject Registration (Admin)"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbackstd 
      Caption         =   "BACK"
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
      TabIndex        =   10
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton cmdsubmitstd 
      Caption         =   "SUBMIT"
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
      Left            =   840
      TabIndex        =   9
      Top             =   4320
      Width           =   3615
   End
   Begin VB.OptionButton optsession3std 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "3 (8.00 pm -  9.00 pm)"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   3120
      Width           =   2535
   End
   Begin VB.OptionButton optsession2std 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "2 (10.00 am -  11.00 am)"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.OptionButton optsession4std 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "4 (9.00 pm -  10.00 pm)"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.OptionButton optsession1std 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "1 (9.00 am -  10.00 am)"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ComboBox cmbdaystd 
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
      ItemData        =   "frmadmintimetablestd.frx":0000
      Left            =   5760
      List            =   "frmadmintimetablestd.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ComboBox cmbsubjectstd 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmadmintimetablestd.frx":005D
      Left            =   2760
      List            =   "frmadmintimetablestd.frx":005F
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      Caption         =   "SESSION"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1935
      Left            =   720
      TabIndex        =   5
      Top             =   2040
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "GENERATE TIMETABLE"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label lbldaystd 
      BackColor       =   &H00404080&
      Caption         =   "DAY"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblsubjectcodestd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SUBJECT CODE"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmadmintimetablestd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strtimetable As String
Dim stdid As String

Private Sub cmbsubjectstd_GotFocus()
    'Make sure click on Microsoft AxtiveX Data object 2.8 Libarary selected
    'Click menu Project, click reference, click Microsoft AxtiveX Data Object 2.8 Libarary, click ok

    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    OpenTuitionDatabase
    OpenLearnTable
    OpenLearnTempTable
    
    stdid = frmadminstudent.txtidstd
    
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
    rst.Open "Select * FROM LearnTemp;", cnn, adOpenStatic
    rst.MoveFirst
    cmbsubjectstd.Clear
        Do
            If rst!std_id = stdid Then
                cmbsubjectstd.AddItem rst![sub_code]
            End If
            rst.MoveNext
        Loop Until rst.EOF
    
End Sub

Private Sub cmdbackstd_Click()
    frmadminstudent.Show
    frmadmintimetablestd.Hide
End Sub

Private Sub cmdsubmitstd_Click()
    OpenTuitionDatabase
    OpenLearnTable
    OpenLearnTempTable
    
        strtimetable = frmadminstudent.txtidstd
        
        rs8.Index = "std_id"
        rs8.Seek "=", strtimetable
        'rs8.Edit
    If rs8!sub_code = cmbsubjectstd Then
        rs5.AddNew
        rs5!std_id = strtimetable
        rs5!sub_code = cmbsubjectstd
        rs5!timetable_day = cmbdaystd
        
        If optsession1std.Value = True Then
            rs5!timetable_session = 1
        ElseIf optsession2std.Value = True Then
            rs5!timetable_session = 2
        ElseIf optsession3std.Value = True Then
            rs5!timetable_session = 3
        ElseIf optsession4std.Value = True Then
            rs5!timetable_session = 4
        End If
        rs5.Update
        MsgBox "Record had been add ", vbOKOnly, "Add Record"
        rs8.Delete
    Else
        MsgBox "Record had not been add ", vbOKOnly, "Add Record"
    End If
    
End Sub

