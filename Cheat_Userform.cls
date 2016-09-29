'=================== CHEAT FORM ==================='
'Created By: Brady Hammond
'Created: 08/18/16
'=================================================='

'Initialize User Form Subroutine:
Private Sub UserForm_Initialize()

    'Hide all advisement buttons
    Cheat_User_Form.AIP_Button.Visible = False
    Cheat_User_Form.Major_Exploration_Button.Visible = False
    Cheat_User_Form.Change_of_Major_Button.Visible = False
    Cheat_User_Form.Initial_Class_Plan_Button.Visible = False
    Cheat_User_Form.Revise_Class_Plan_Button.Visible = False
    Cheat_User_Form.Graduation_Clearance_Button.Visible = False
    Cheat_User_Form.One_Semester_Planning_Button.Visible = False
    Cheat_User_Form.Over_Hours_Change_of_Major_Button.Visible = False
    Cheat_User_Form.Internship_Orientation_Button.Visible = False
    Cheat_User_Form.Internship_Exit_Interview_Button.Visible = False
    Cheat_User_Form.GE_Questions_Button.Visible = False
    Cheat_User_Form.GE_Substitution_Button.Visible = False
    Cheat_User_Form.Major_Advising_Button.Visible = False
    
    'Show all subject buttons
    Cheat_User_Form.Biology_Department_Button.Visible = True
    Cheat_User_Form.Biological_Science_Education_Button.Visible = True
    Cheat_User_Form.Athletic_Training_Button.Visible = True
    Cheat_User_Form.Exercise_and_Wellness_Button.Visible = True
    Cheat_User_Form.Exercise_Science_I_Button.Visible = True
    Cheat_User_Form.Exercise_Science_II_Button.Visible = True
    Cheat_User_Form.Health_Science_Department_Button.Visible = True
    Cheat_User_Form.School_House_Education_Button.Visible = True
    Cheat_User_Form.MMBIO_Department_Button.Visible = True
    Cheat_User_Form.Food_Science_Button.Visible = True
    Cheat_User_Form.Nutritional_Science_Button.Visible = True
    Cheat_User_Form.Dietetics_Button.Visible = True
    Cheat_User_Form.PDBIO_Department_Button.Visible = True
    Cheat_User_Form.PWS_Department_Button.Visible = True
    
    'Reset the variable holder
    Cheat_User_Form.Variable_Holder.Visible = False
    Cheat_User_Form.Variable_Holder.Caption = ""
    
End Sub
  
'Cancel Button Subroutine:
Private Sub Cancel_Button_Click()
    
    'Unload the user form
    Unload Me
    
End Sub

'Biology Department Button Click Subroutine:
Private Sub Biology_Department_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Biology Department,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Biological Science Education Button Click Subroutine:
Private Sub Biological_Science_Education_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Biological Science Education,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Athletic Training Button Click Subroutine:
Private Sub Athletic_Training_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Athletic Training,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Exercise and Wellness Button Click Subroutine:
Private Sub Exercise_and_Wellness_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Exercise and Wellness,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Exercise Science I Button Click Subroutine:
Private Sub Exercise_Science_I_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Exercise Science I,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Exercise Science II Button Click Subroutine:
Private Sub Exercise_Science_II_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Exercise Science II,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Health Science Department Button Click Subroutine:
Private Sub Health_Science_Department_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Health Science,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'School House Education Button Click Subroutine:
Private Sub School_House_Education_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "School House Education,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'MMBIO Department Button Click Subroutine:
Private Sub MMBIO_Department_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "MMBIO Department,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Food Science Button Click Subroutine:
Private Sub Food_Science_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Food Science,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Nutritional Science Button Click Subroutine:
Private Sub Nutritional_Science_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Nutritional Science,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'Dietetics Button Click Subroutine:
Private Sub Dietetics_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "Dietetics,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'PDBIO Department Button Click Subroutine:
Private Sub PDBIO_Department_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "PDBIO Department,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'PWS Department Button Click Subroutine:
Private Sub PWS_Department_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = "PWS Department,"
    
    'Flip buttons
    Call Flip_Buttons
    
End Sub

'AIP Button Subroutine:
Private Sub AIP_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "AIP"
    
    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Major Exploration Button Subroutine:
Private Sub Major_Exploration_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Major Exploration"
    
    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Change of Major Button Subroutine:
Private Sub Change_of_Major_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Change of Major"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Initial Class Plan Button Subroutine:
Private Sub Initial_Class_Plan_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Initial Class Plan"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Revise Class Plan Button Subroutine:
Private Sub Revise_Class_Plan_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Revise Class Plan"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Graduation Clearance Button Subroutine:
Private Sub Graduation_Clearance_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Graduation Clearance"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'One Semester Planning Button Subroutine:
Private Sub One_Semester_Planning_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "One Semester Planning"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Over Hours Change of Major Button Subroutine:
Private Sub Over_Hours_Change_of_Major_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Over Hours Change of Major"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Internship Orientation Button Subroutine:
Private Sub Internship_Orientation_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Internship Orientation"
    
    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Internship Exit Interview Button Subroutine:
Private Sub Internship_Exit_Interview_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Internship Exit Interview"
    
    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'GE Questions Button Subroutine:
Private Sub GE_Questions_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "GE Questions"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'GE Substitution Button Subroutine:
Private Sub GE_Substitution_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "GE Substitution"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Major Advising Button Subroutine:
Private Sub Major_Advising_Button_Click()

    'Record button
    Cheat_User_Form.Variable_Holder.Caption = Cheat_User_Form.Variable_Holder.Caption & "Major Advising"

    'Cross check the buttons
    Call Button_Cross_Check
    
    'Unload the user form
    Unload Me
    
End Sub

'Flip Buttons Subroutine
Private Sub Flip_Buttons()

    'Show all advisement buttons
    Cheat_User_Form.AIP_Button.Visible = True
    Cheat_User_Form.Major_Exploration_Button.Visible = True
    Cheat_User_Form.Change_of_Major_Button.Visible = True
    Cheat_User_Form.Initial_Class_Plan_Button.Visible = True
    Cheat_User_Form.Revise_Class_Plan_Button.Visible = True
    Cheat_User_Form.Graduation_Clearance_Button.Visible = True
    Cheat_User_Form.One_Semester_Planning_Button.Visible = True
    Cheat_User_Form.Over_Hours_Change_of_Major_Button.Visible = True
    Cheat_User_Form.Internship_Orientation_Button.Visible = True
    Cheat_User_Form.Internship_Exit_Interview_Button.Visible = True
    Cheat_User_Form.GE_Questions_Button.Visible = True
    Cheat_User_Form.GE_Substitution_Button.Visible = True
    Cheat_User_Form.Major_Advising_Button.Visible = True
    
    'Hide all subject buttons
    Cheat_User_Form.Biology_Department_Button.Visible = False
    Cheat_User_Form.Biological_Science_Education_Button.Visible = False
    Cheat_User_Form.Athletic_Training_Button.Visible = False
    Cheat_User_Form.Exercise_and_Wellness_Button.Visible = False
    Cheat_User_Form.Exercise_Science_I_Button.Visible = False
    Cheat_User_Form.Exercise_Science_II_Button.Visible = False
    Cheat_User_Form.Health_Science_Department_Button.Visible = False
    Cheat_User_Form.School_House_Education_Button.Visible = False
    Cheat_User_Form.MMBIO_Department_Button.Visible = False
    Cheat_User_Form.Food_Science_Button.Visible = False
    Cheat_User_Form.Nutritional_Science_Button.Visible = False
    Cheat_User_Form.Dietetics_Button.Visible = False
    Cheat_User_Form.PDBIO_Department_Button.Visible = False
    Cheat_User_Form.PWS_Department_Button.Visible = False
    
End Sub

'Button Cross Check Subroutine:
Private Sub Button_Cross_Check()

    'Declare needed variables
    Dim Split_Array As Variant
    Dim Advisement As String
    Dim Subject As String
    
    Split_Array = Split(Cheat_User_Form.Variable_Holder.Caption, ",")
    Subject = Split_Array(0)
    Advisement = Split_Array(1)
    
    If Subject = "Biology Department" Then
    
        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Biological Science Education" Then
    
        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Athletic Training" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "Marcia Richards", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Exercise and Wellness" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Exercise Science I" Then
    
        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Exercise Science II" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
                
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Health Science" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "School House Education" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "Students CANNOT change their major to School House Education", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "MMBIO Department" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Food Science" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Nutritional Science" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "Dietetics" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "PDBIO Department" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    ElseIf Subject = "PWS Department" Then

        If Advisement = "AIP" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Major Exploration" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Initial Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Revise Class Plan" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Graduation Clearance" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "One Semester Planning" Then
            
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Over Hours Change of Major" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Internship Orientation" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
        
        ElseIf Advisement = "Internship Exit Interview" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Questions" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "GE Substitution" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        ElseIf Advisement = "Major Advising" Then
        
            'Output the correct advisor
            MsgBox "XXXXX", vbInformation, "Advisor"
            
        Else
        
            'Output an error message
            MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
        End If
        
    Else
    
        'Output an error message
        MsgBox "There is a programming issue. Please inform the current programmer of this message", vbCritical, "Programming Error"
        
    End If
    
End Sub
