VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DueDatesForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4784
   ClientLeft      =   -40
   ClientTop       =   -136
   ClientWidth     =   6928
   OleObjectBlob   =   "DueDatesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DueDatesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    With Me.CBCourse
        .AddItem "MATH 115"
        .AddItem "MATH 116"
        .AddItem "MSE 121"
        .AddItem "MSE 100"
        .AddItem "CHE 102"
        .AddItem "GENE 119"
    End With
    
    With Me.CBType
        .AddItem "Project"
        .AddItem "Test"
        .AddItem "Quiz"
        .AddItem "Exam"
        .AddItem "Assignment"
    End With
    
    With Me.CBStatus
        .AddItem "COMPLETED"
        .AddItem "NOT STARTED"
        .AddItem "IN PROGRESS"
    End With
    
    With Me.CBPriority
        .AddItem "HIGH"
        .AddItem "MEDIUM"
        .AddItem "LOW"
    End With
End Sub
Private Sub SubmitButton_Click()
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Due Dates")
    
    'Submission for Assignment Information
    intRow = 3
    If (nametxt.Value <> "") Then
        If (CBCourse.Value <> "" And CBType.Value <> "") Then
            Do While (ws.Cells(intRow, "A") <> "")
                intRow = intRow + 1
            Loop
            
            ws.Cells(intRow, "A") = nametxt.Value
            ws.Cells(intRow, "B") = CBCourse.Value
            ws.Cells(intRow, "C") = CBType.Value
        Else
            MsgBox "Please enter a course and assignment type."
            nametxt.Value = ""
            CBCourse.Value = ""
            CBType.Value = ""
     
            daytxt.Value = ""
            monthtxt.Value = ""
            yeartxt.Value = ""
            CBPriority.Value = ""
            CBStatus.Value = ""
            DueDatesForm.Hide
            Exit Sub
        End If
    Else
        MsgBox "Please enter a name for the assignment."
    End If
    
    'Submission for Assignment Status
    intRow = 3
    
    If (CBStatus.Value <> "" And CBPriority.Value <> "") Then
        If (daytxt.Value <> "" And yeartxt.Value <> "" And monthtxt <> "" And IsNumeric(daytxt.Value) And IsNumeric(yeartxt.Value) And IsNumeric(monthtxt.Value)) Then
            Do While (ws.Cells(intRow, "D") <> "")
                intRow = intRow + 1
            Loop
            
            ws.Cells(intRow, "D") = yeartxt.Value + "-" + monthtxt.Value + "-" + daytxt.Value
            ws.Cells(intRow, "D").NumberFormat = "yyyy-mm-dd;@"
            ws.Cells(intRow, "E") = CBStatus.Value
            ws.Cells(intRow, "F") = CBPriority.Value
        Else
            MsgBox "Please enter a valid numeric day, month, and year."
            nametxt.Value = ""
            CBCourse.Value = ""
            CBType.Value = ""
     
            daytxt.Value = ""
            monthtxt.Value = ""
            yeartxt.Value = ""
            CBPriority.Value = ""
            CBStatus.Value = ""
            DueDatesForm.Hide
            Exit Sub
        End If
    Else
        MsgBox "Please enter a status and priority"
        nametxt.Value = ""
        CBCourse.Value = ""
        CBType.Value = ""
                 
        daytxt.Value = ""
        monthtxt.Value = ""
        yeartxt.Value = ""
        CBPriority.Value = ""
        CBStatus.Value = ""
        DueDatesForm.Hide
        Exit Sub
    End If
        
    nametxt.Value = ""
    CBCourse.Value = ""
    CBType.Value = ""
     
    daytxt.Value = ""
    monthtxt.Value = ""
    yeartxt.Value = ""
    CBPriority.Value = ""
    CBStatus.Value = ""
    DueDatesForm.Hide
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub





