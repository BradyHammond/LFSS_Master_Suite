'=================== SCANNING FORM ==================='
'Created By: Brady Hammond
'Created: 05/09/16
'====================================================='

'Initialize User Form Subroutine:
Private Sub UserForm_Initialize()
    
    'Set the options for the combo box
    Name_Files_Combo.List = Array("Plan", "ProgressReport", "Waiver", "Transfer", "Substitution", "AcademicImprovementPlan", "Email", "ChangeOfMajor", "Minor", "Other (Custom)")
    
    'Set the focus on the combo box
    Name_Files_Combo.SetFocus
        
End Sub

'Skip Button Subroutine:
Private Sub Skip_Button_Click()

    'Get the row of the file to be scanned
    Dim Row
    Row = Scanning_User_Form.Row.Caption
    
    'Create a status variable to show if a scan was skipped
    Dim Status
    Status = "Skipped"
    
    'Write "Skipped" in the status column and color the cell red
    ActiveWorkbook.ActiveSheet.Range("E" & Row).Value = Status
    ActiveWorkbook.ActiveSheet.Range("E" & Row).Interior.Color = RGB(255, 0, 0)
    
    'Load the next page
    Call Load_Next_PDF(Scanning_User_Form.Row.Caption, ActiveWorkbook.ActiveSheet, "", Scanning_User_Form.Files_To_Be_Scanned.Caption)
    
    'Set the focus to the Student_ID_Box
    Scanning_User_Form.Name_Files_Combo.SetFocus
    
End Sub

'Submit Button (Enter) Subroutine:
Private Sub Submit_Button_Enter()

    'Click the submit button
    Call Submit_Button_Click
    
End Sub

'Submit Button (Click) Subroutine:
Private Sub Submit_Button_Click()

    'Declare needed variables
    Dim Student_ID
    Dim Worksheet As Worksheet
    
    'Define Student_ID, File, and Worksheet
    Student_ID = Scanning_User_Form.Student_ID_Box.Value
    Set Worksheet = ActiveWorkbook.ActiveSheet
    
    'Check if the file has been named
    If File_Name = "" Then
    
        'If not, output an error message and exit the subroutine
        MsgBox "Please name this file.", vbInformation, "No file name given"
        Exit Sub
        
    End If
    
    'Remove dashes from the ID number
    Student_ID = Replace(Student_ID, "-", "")
    
    'Remove spaces from the ID number
    Student_ID = Replace(Student_ID, " ", "")
    
    'Check if the ID is a number
    If IsNumeric(Student_ID) Then
    
        'Check if the ID number is the correct lenght
        If Not Len(Student_ID) = 9 Then
            
            'If not, output an error message, reset the focus, and exit the subroutine
            MsgBox "The ID number entered was not 9 characters long. Please enter a valid ID number.", vbInformation, "Invalid ID Number"
            Scanning_User_Form.Student_ID_Box.SetFocus
            Exit Sub
        
        Else
            
            'Save the ID number and search for it on the AIM system
            Scanning_User_Form.Student_ID_Box.Value = Student_ID
            Call Search_ID_Number(Student_ID, Worksheet)
            Call Load_Next_PDF(Scanning_User_Form.Row.Caption, ActiveWorkbook.ActiveSheet, "", Scanning_User_Form.Files_To_Be_Scanned.Caption)
            
        End If
        
    Else
        
        'Search the AIM system for the netID
        Call Search_ID_Number(Student_ID, Worksheet)
        Call Load_Next_PDF(Scanning_User_Form.Row.Caption, ActiveWorkbook.ActiveSheet, "", Scanning_User_Form.Files_To_Be_Scanned.Caption)
    
    End If
    
End Sub

'Cancel Button (Click) Subroutine:
Private Sub Cancel_Button_Click()
    
    'Declare needed variables
    Dim Computer As String
    Dim Object_WMI_Service
    Dim Items_Collection
    Dim Object_Item
    Dim Response
    
    'Define needed variables
    Computer = "."
    Set Object_WMI_Service = GetObject("winmgmts:\\" & Computer & "\root\CIMV2")
    Set Items_Collection = Object_WMI_Service.ExecQuery("SELECT * FROM Win32_Process", , 48)
    
    'Search through the running programs
    For Each Object_Item In Items_Collection
    
        'Check for Adobe Acrobat
        If Object_Item.Name = "Acrobat.exe" Then
            
            'If found close it
            Object_Item.Terminate
            
        End If
        
    Next
    
    'Close the form
    Unload Scanning_User_Form
    
    'Ask if the user wants to close the workbook
    Response = MsgBox("The scanning process was not completed. Would you like to close this excel sheet now?", vbInformation + vbYesNo, "Incomplete Scanning Process")
    
    'Check if the user said yes
    If Response = vbYes Then
    
        'If so close the workbook
        ActiveWorkbook.Close SaveChanges:=False
        
    End If
    
End Sub

