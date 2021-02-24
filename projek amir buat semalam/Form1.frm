VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton posttest 
      Caption         =   "Post"
      Height          =   1575
      Left            =   4320
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.PictureBox Picloop 
      Height          =   6135
      Left            =   360
      ScaleHeight     =   6075
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton pretest 
      Caption         =   "Pre"
      Height          =   1335
      Left            =   4320
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub pretest_Click()
Dim Q, K As Integer
    Picloop.Cls
    Q = InputBox("Enter number", "Number", 5)
    Picloop.Print Tab(15); “Q”; Spc(15); “K”
    Picloop.Print
    intLoop = 10
    Do While Q < 10
        K = (2 * Q) - 1
        Q = Q + 1
        Picloop.Print Tab(15); “Q”; Spc(15); “K”
    Loop
End Sub


Private Sub posttest_Click()
Dim Q, K As Integer
    Picloop.Cls
    Q = InputBox("Enter number", "Number", 5)
    Picloop.Print Tab(15); “Q”; Spc(15); “K”
    Picloop.Print
    intLoop = 10
    Do
        K = (2 * Q) - 1
        Q = Q + 1
        Picloop.Print Tab(15); “Q”; Spc(15); “K”
    Loop While Q < 10
End Sub
