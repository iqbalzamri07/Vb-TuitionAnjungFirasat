Attribute VB_Name = "Moduleteacher"
Option Explicit
 
Public dbtuition  As Database   'open Tuition database
Public rs1 As Recordset 'Student
Public rs2 As Recordset 'Teacher
Public rs3 As Recordset 'Admin
Public rs4 As Recordset 'Subject
Public rs5 As Recordset 'Learn
Public rs6 As Recordset 'Deal
Public rs7 As Recordset 'Teach
Public rs8 As Recordset 'LearnTemp
Public rs9 As Recordset 'TeachTemp

'------------------------------------------------------------------------
' Procedure to open the Tuition database
'------------------------------------------------------------------------
Public Sub OpenTuitionDatabase()
 
    Set dbtuition = OpenDatabase(GetAppPath() & "Tuition.mdb")
 
End Sub

'----------------------------------------------------------------------
'Procedure to open Student table
'----------------------------------------------------------------------
Public Sub OpenStudentTable()
  
    Set rs1 = dbtuition.OpenRecordset("Student")

End Sub
 
'----------------------------------------------------------------------
'Procedure to open Teacher table
'----------------------------------------------------------------------
Public Sub OpenTeacherTable()
  
    Set rs2 = dbtuition.OpenRecordset("Teacher")

End Sub
 
'----------------------------------------------------------------------
'Procedure to open Admin table
'----------------------------------------------------------------------
Public Sub OpenAdminTable()
  
    Set rs3 = dbtuition.OpenRecordset("Admin")

End Sub

'----------------------------------------------------------------------
'Procedure to open Subject table
'----------------------------------------------------------------------
Public Sub OpenSubjectTable()
  
    Set rs4 = dbtuition.OpenRecordset("Subject")

End Sub

'----------------------------------------------------------------------
'Procedure to open Learn table
'----------------------------------------------------------------------
Public Sub OpenLearnTable()
  
    Set rs5 = dbtuition.OpenRecordset("Learn")

End Sub

'----------------------------------------------------------------------
'Procedure to open Deal table
'----------------------------------------------------------------------
Public Sub OpenDealTable()
  
    Set rs6 = dbtuition.OpenRecordset("Deal")

End Sub

'----------------------------------------------------------------------
'Procedure to open Teach table
'----------------------------------------------------------------------
Public Sub OpenTeachTable()
  
    Set rs7 = dbtuition.OpenRecordset("Teach")

End Sub

'----------------------------------------------------------------------
'Procedure to open LearnTemp table
'----------------------------------------------------------------------
Public Sub OpenLearnTempTable()
  
    Set rs8 = dbtuition.OpenRecordset("LearnTemp")

End Sub
 
'----------------------------------------------------------------------
'Procedure to open TeachTemp table
'----------------------------------------------------------------------
Public Sub OpenTeachTempTable()
  
    Set rs9 = dbtuition.OpenRecordset("TeachTemp")

End Sub
'------------------------------------------------------------------------
'Procedure to close Tuition database
'------------------------------------------------------------------------
Public Sub CloseEmpDatabase()

    dbtuition.Close
    Set dbtuition = Nothing
 
End Sub

'------------------------------------------------------------------------
Public Sub CenterForm(pobjForm As Form)
'------------------------------------------------------------------------
 
    With pobjForm
    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2
End With
End Sub
 
'------------------------------------------------------------------------
Public Function GetAppPath() As String
'------------------------------------------------------------------------
 
    GetAppPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")
 
End Function


