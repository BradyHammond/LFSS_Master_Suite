'=================== EMAIL FORM ==================='
'Created By: Brady Hammond
'Created: 05/25/16
'=================================================='

'Initialize User Form Subroutine:
Private Sub UserForm_Initialize()
    
    'Declare needed variables
    Dim Day As String
    Dim Month As String
    Dim Year As String
    
    'Define the day month and year
    Day = DatePart("d", Date)
    Month = DatePart("m", Date)
    Year = DatePart("yyyy", Date)
    
    'Check if the day has two characters in it
    If Len(Day) < 2 Then
        
        'If not, add a zero in the front
        Day = "0" & Day
        
    End If
    
    'Check if the month has two characters in it
    If Len(Month) < 2 Then
    
        'If not, add a zero in the front
        Month = "0" & Month
        
    End If
    
    'Fill Text Fields
    Email_User_Form.Specific_Date_Day_Input.Value = Day
    Email_User_Form.Specific_Date_Month_Input.Value = Month
    Email_User_Form.Specific_Date_Year_Input.Value = Year
    
    Email_User_Form.End_Date_Day_Input.Value = Day
    Email_User_Form.End_Date_Month_Input.Value = Month
    Email_User_Form.End_Date_Year_Input.Value = Year
    
End Sub

'Cancel Button Click Subroutine:
Private Sub Cancel_Button_Click()

    'Unload the user form
    Unload Email_User_Form
    
End Sub

'Tomorrow's Appointments Button Click Subroutine:
Private Sub Tomorrow_Button_Click()

    'Get the calendar data for tomorrow
    Call Get_Calendar_Data(Date + 1)
    
    'Send the emails
    Call Send_Emails
    
End Sub

'Today's Appointments Button Click Subroutine:
Private Sub Today_Button_Click()
    
    'Get the calendar data for today
    Call Get_Calendar_Data(Date)
    
    'Send the emails
    Call Send_Emails
    
End Sub

'Monday's Appointments Button Click Subroutine:
Private Sub Monday_Button_Click()
    
    'Declare needed variables
    Dim Day_Of_The_Week As Integer
    
    'Find which day of the week it is
    Day_Of_The_Week = Weekday(Date)
    
    'Check if it's Sunday
    If Day_Of_The_Week = 1 Then

        'If so, get the calendar data for Monday
        Call Get_Calendar_Data(Date + 1)
        
        'Send the emails
        Call Send_Emails
    
    'Check if it's already Monday
    ElseIf Day_Of_The_Week = 2 Then
    
        'If so, output an error message and exit the subroutine
        MsgBox "It is already Monday. If you wish to send out the emails for today click the Today's Appointments button.", vbInformation, "Already Monday"
        Exit Sub
        
    'Check if it's Tuesday
    ElseIf Day_Of_The_Week = 3 Then

        'If so, get the calendar data for Monday
        Call Get_Calendar_Data(Date + 6)
        
        'Send the emails
        Call Send_Emails
    
    'Check if it's Wednesday
    ElseIf Day_Of_The_Week = 4 Then
        
        'If so, get the calendar data for Monday
        Call Get_Calendar_Data(Date + 5)
        
        'Send the emails
        Call Send_Emails
        
    'Check if it's Thursday
    ElseIf Day_Of_The_Week = 5 Then

        'If so, get the calendar data for Monday
        Call Get_Calendar_Data(Date + 4)
        
        'Send the emails
        Call Send_Emails
    
    'Check if it's Friday
    ElseIf Day_Of_The_Week = 6 Then
    
        'If so, get the calendar data for Monday
        Call Get_Calendar_Data(Date + 3)
        
        'Send the emails
        Call Send_Emails
        
    'Check if it's Saturday
    ElseIf Day_Of_The_Week = 7 Then
    
        'If so, get the calendar data for Monday
        Call Get_Calendar_Data(Date + 2)
        
        'Send the emails
        Call Send_Emails
        
    End If
    
End Sub

'Specific Date/Dates Button Click Subroutine:
Private Sub Specific_Date_Button_Click()

    'Check if the input is comprised of numbers
    If IsNumeric(Email_User_Form.Specific_Date_Month_Input.Value) = True And IsNumeric(Email_User_Form.Specific_Date_Day_Input.Value) = True And IsNumeric(Email_User_Form.Specific_Date_Year_Input.Value) = True And IsNumeric(Email_User_Form.End_Date_Month_Input.Value) = True And IsNumeric(Email_User_Form.End_Date_Day_Input.Value) = True And IsNumeric(Email_User_Form.End_Date_Year_Input.Value) Then
        
        'Check that the months are correct
        If Email_User_Form.Specific_Date_Month_Input.Value <= 12 And Email_User_Form.End_Date_Month_Input.Value <= 12 Then
        
        
            'Check if the date is valid
            If IsDate(Email_User_Form.Specific_Date_Month_Input.Value & "/" & Email_User_Form.Specific_Date_Day_Input.Value & "/" & Email_User_Form.Specific_Date_Year_Input.Value) = True And IsDate(Email_User_Form.End_Date_Month_Input.Value & "/" & Email_User_Form.End_Date_Day_Input.Value & "/" & Email_User_Form.End_Date_Year_Input.Value) = True Then
        
                'If so, get the calendar data for the listed date(s)
                Call Get_Calendar_Data(Email_User_Form.Specific_Date_Month_Input.Value & "/" & Email_User_Form.Specific_Date_Day_Input.Value & "/" & Email_User_Form.Specific_Date_Year_Input.Value, Email_User_Form.End_Date_Month_Input.Value & "/" & Email_User_Form.End_Date_Day_Input.Value & "/" & Email_User_Form.End_Date_Year_Input.Value)
            
                'Send the emails
                Call Send_Emails
        
            Else
        
                'Otherwise output an error message
                MsgBox "There is a problem with one of the dates you have given. Please check them and try again", vbInformation, "Invalid Date"
        
            End If
            
        Else
        
            'Otherwise output an error message
            MsgBox "There is a problem with one of the dates you have given. Please check them and try again", vbInformation, "Invalid Date"
        
        End If
        
    Else
        
        'Otherwise output an error message
        MsgBox "There is a problem with one of the dates you have given. Please check them and try again", vbInformation, "Invalid Date"
        
    End If
    
End Sub

'Send Survey Button Click Subroutine:
Private Sub Survey_Button_Click()

    'Declare needed variables
    Dim Day_Of_The_Week As Integer
    
    'Find which day of the week it is
    Day_Of_The_Week = Weekday(Date)
    
    'Check if it is Sunday
    If Day_Of_The_Week = 1 Then
    
        'If so, get the calendar data from Friday
        Call Get_Calendar_Data(Date - 2)
        
        'Send the survey
        Call Send_Survey
    
    'Check if it is Monday
    ElseIf Day_Of_The_Week = 2 Then
    
        'If so, get the calendar data from Friday
        Call Get_Calendar_Data(Date - 3)
        
        'Send the survey
        Call Send_Survey
    
    Else
    
        'Get the calendar data from the previous day
        Call Get_Calendar_Data(Date - 1)
        
        'Send the survey
        Call Send_Survey
    
    End If
    
End Sub
