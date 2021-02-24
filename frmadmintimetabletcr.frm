VERSION 5.00
Begin VB.Form frmadmintimetabletcr 
   BackColor       =   &H00404080&
   Caption         =   "Teacher Subject Registration (Admin)"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbacktcr 
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
      Top             =   4440
      Width           =   3375
   End
   Begin VB.ComboBox cmbsubjecttcr 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmadmintimetabletcr.frx":0000
      Left            =   2760
      List            =   "frmadmintimetabletcr.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ComboBox cmbdaytcr 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmadmintimetabletcr.frx":0004
      Left            =   5760
      List            =   "frmadmintimetabletcr.frx":001D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.OptionButton optsession1tcr 
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
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
   Begin VB.OptionButton optsession4tcr 
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
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   3360
      Width           =   2535
   End
   Begin VB.OptionButton optsession2tcr 
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
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2640
      Width           =   2535
   End
   Begin VB.OptionButton optsession3tcr 
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
      TabIndex        =   1
      Top             =   3240
      Width           =   2535
   End
   Begin VB.CommandButton cmdsubmittcr 
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
      TabIndex        =   0
      Top             =   4440
      Width           =   3615
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
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   720
      TabIndex        =   7
      Top             =   2160
      Width           =   7455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   1200
      TabIndex        =   11
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label lblsubjectcodetcr 
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lbldaytcr 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "frmadmintimetabletcr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strtimetable As String
Dim tcrid As String
Dim countt As Integer


Private Sub cmbsubjecttcr_GotFocus()
    'Make sure click on Microsoft AxtiveX Data object 2.8 Libarary selected
    'Click menu Project, click reference, click Microsoft AxtiveX Data Object 2.8 Libarary, click ok
    countt = 0
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    OpenTuitionDatabase
    OpenTeachTable
    OpenTeachTempTable
    tcrid = frmadminteacher.txtidtcr
    
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Tuition.mdb"
    rst.Open "Select * FROM TeachTemp;", cnn, adOpenStatic
    rst.MoveFirst
    cmbsubjecttcr.Clear
        Do
            If rst!tcr_id = tcrid Then
                cmbsubjecttcr.AddItem rst![sub_code]
                countt = countt + 1
            End If
            rst.MoveNext
        Loop Until rst.EOF
    
End Sub

Private Sub cmdbacktcr_Click()
    frmadminteacher.Show
    frmadmintimetabletcr.Hide
End Sub

Private Sub cmdsubmittcr_Click()
    OpenTuitionDatabase
    OpenTeachTable
    OpenTeachTempTable
    
        strtimetable = frmadminteacher.txtidtcr
        
        rs9.Index = "tcr_id"
        rs9.Seek "=", strtimetable
        'rs9.Edit
    If rs9!sub_code = cmbsubjecttcr Then
        rs7.AddNew
        rs7!tcr_id = strtimetable
        rs7!sub_code = cmbsubjecttcr
        rs7!timetable_day = cmbdaytcr
        
        If optsession1tcr.Value = True Then
            rs7!timetable_session = 1
        ElseIf optsession2tcr.Value = True Then
            rs7!timetable_session = 2
        ElseIf optsession3tcr.Value = True Then
            rs7!timetable_session = 3
        ElseIf optsession4tcr.Value = True Then
            rs7!timetable_session = 4
        End If
        rs7.Update
        MsgBox "Record had been add ", vbOKOnly, "Add Record"
        rs9.Delete
    Else
        MsgBox "Record had not been add ", vbOKOnly, "Add Record"
    End If
    
End Sub
