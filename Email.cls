'========================= EMAILS ========================='
'Created By: Brady Hammond
'Created: 05/23/16
'=========================================================='

'Emails Subroutine:
Sub EMAIL_PROGRAM()

    'Show the user form
    Email_User_Form.Show
    
End Sub

'Get Calandar Data Subroutine:
Sub Get_Calendar_Data(Starting_Date As Date, Optional Ending_Date As Date)
    
    Dim Recipients(0 To 5) As String
    Dim Recipients_Index As Integer
    Dim Workbook As Excel.Workbook
    Dim Starting_Cell As Excel.Range
    Dim Regex_Pattern As Object
    Dim Excel_Range As Range
    Dim Excel_Cell As Range
    Dim Row As Long
    
    'Check if the starting date is valid
    If IsDate(Starting_Date) = False Then
    
        'If not, exit the subroutine
        Exit Sub
        
    End If
    
    'Check if an ending date was used (the ending date will be 12:00:00 AM if not explicitly given)
    If Ending_Date = "12:00:00 AM" Then
    
        'If so, set the ending date to the starting date
        Ending_Date = Starting_Date
    
    End If
    
    'Check if the ending date is valid
    If IsDate(Ending_Date) = False Then
    
        'If not, exit the subroutine
        Exit Sub
        
    End If
    
    'Check if the ending date is before the starting date
    If Ending_Date < Starting_Date Then
    
        'Output an error message and exit the subroutine
        MsgBox "The ending date cannont be before the starting date. Please check the dates and try again", vbInformation, "Invalid Dates"
        Exit Sub
        
    End If
    
    'Check if the requested time span is a long time
    If Ending_Date - Starting_Date > 28 Then
    
        'If so, ask if the user really wants that much info
        If MsgBox("This could take some time. Do you wish to continue anyway?", vbInformation + vbYesNo, "Continue") = vbNo Then
        
            'If so, exit the subroutine:
            Exit Sub
            
        End If
        
    End If
    
    'Output a wait message
    MsgBox "Please wait a moment while the program compiles the calendar data.", vbInformation, "Compilation in Progress"
    
    'Create the regex script
    Set Regex_Pattern = CreateObject("vbscript.regexp")
    
    'Define the regex pattern
    Regex_Pattern.Pattern = ("\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b")
    
    'Make the pattern global
    Regex_Pattern.Global = True
    
    'Disable case sensitivity
    Regex_Pattern.ignoreCase = True
    
    'Define the recipients
    Recipients(0) = "XXXX"
    Recipients(1) = "XXXX"
    Recipients(2) = "XXXX"
    Recipients(3) = "XXXX"
    Recipients(4) = "XXXX"
    Recipients(5) = "XXXX"
    
    'Create the excel workbook
    Set Workbook = Excel.Workbooks.Add
    
    'Define the start cell
    Set Starting_Cell = Workbook.Sheets(1).Range("A1")
    
    'Turn of screen updating
    Application.ScreenUpdating = False
    
    'Fill in the headers on the worksheet
    With Starting_Cell
    
        .Offset(0, 0).Value = "Subject"
        .Offset(0, 1).Value = "Starting Date"
        .Offset(0, 2).Value = "Starting Time"
        .Offset(0, 3).Value = "Ending Time"
        .Offset(0, 4).Value = "Text"
        .Offset(0, 5).Value = "Phone"
        .Offset(0, 6).Value = "Email"
        .Offset(0, 7).Value = "Advisor"
        .Offset(0, 8).Value = "Appointment Type"
    
    End With
    
    'Loop through each recipient
    For Recipients_Index = 0 To 5
    
        Call Find_Appointments(Recipients(Recipients_Index), Starting_Cell, Starting_Date, Ending_Date)
        
    Next Recipients_Index
    
    'Set the data range
    Set Excel_Range = Range("H2:H" & Range("A" & Rows.Count).End(xlUp).Row)
    
    'Loop through each cell in the data range
    For Each Cell In Excel_Range
    
        'Activate the cell
        Cell.Activate
        
        'Check if the cell has Jeanne Gubler in it
        If Cell.Value = "Jeanne Gubler" Then
            
            'If so, delete the row and shift the next row up
            Range("A" & ActiveCell.Row & ":J" & ActiveCell.Row).Select
            Selection.Delete Shift:=xlUp
            ActiveCell.Offset(-1, 0).Activate
            
        End If
    
    Next Cell
    
    'Reset the data range
    Set Excel_Range = Range("A2:A" & Range("A" & Rows.Count).End(xlUp).Row)
    
    'Loop through each cell in the data range
    For Each Cell In Excel_Range
    
        'Activate the cell
        Cell.Activate
        
        'Check if the cell has Jeanne Gubler in it
        If Cell.Value = "" Then
            
            'If so, delete the row and shift the next row up
            Range("A" & ActiveCell.Row & ":J" & ActiveCell.Row).Select
            Selection.Delete Shift:=xlUp
            ActiveCell.Offset(-1, 0).Activate
            
        End If
    
    Next Cell
    
    'Reset the data range
    Set Excel_Range = Range("A1:I" & Range("A" & Rows.Count).End(xlUp).Row)
    
    'Set the cell to sort by
    Set Excel_Cell = Range("E1")
    
    'Sort by whether or not the student wanted a text
    Excel_Range.Sort Key1:=Excel_Cell, Order1:=xlDescending, Header:=xlYes
    
    'Rest the data range
    Set Excel_Range = Range("A1:I" & Range("A" & Rows.Count).End(xlUp).Row)
    
    'Create borders around the data
    Excel_Range.Borders(xlInsideVertical).Weight = xlThin
    Excel_Range.Borders(xlEdgeBottom).Weight = xlThin
    Excel_Range.Borders(xlEdgeRight).Weight = xlThin
    Excel_Range.Borders(xlEdgeLeft).Weight = xlThin
    
    'Autofit the data
    Workbook.Sheets(1).Columns("A:I").AutoFit
    
    'Reset the data range
    Set Excel_Range = Range("A1:I1")
    
    'Set up the headers
    With Excel_Range
        
        .Interior.Color = RGB(100, 100, 100)
        .Font.Color = RGB(255, 255, 255)
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeRight).Weight = xlThick
        .Borders(xlEdgeLeft).Weight = xlThick
        .Borders(xlEdgeTop).Weight = xlThick
        
    End With
    
    'Reset the data range
    Set Excel_Range = Range("A2:A" & Range("A" & Rows.Count).End(xlUp).Row)
    
    'Loop through each row
    For Each Cell In Excel_Range
        
        'Activate the cell and save the row
        Cell.Activate
        Row = ActiveCell.Row
        
        'Check if it's an even row
        If Row Mod 2 = 0 Then
            
            'If so, color the row dark grey
            Workbook.Sheets(1).Range("A" & Row & ":I" & Row).Interior.Color = RGB(200, 200, 200)
            
        Else
            
            'Otherwise color the row light grey
            Workbook.Sheets(1).Range("A" & Row & ":I" & Row).Interior.Color = RGB(225, 225, 225)
        
        End If
    
    Next Cell
    
    'Reset the data range
    Set Excel_Range = Range("E2:E" & Range("A" & Rows.Count).End(xlUp).Row)
    
    'Loop through each cell in column E
    For Each Cell In Excel_Range
    
        'Activate the cell
        Cell.Activate
        
        'Check if it has yes in it
        If Cell.Value = "Yes" Then
            
            'If so, color the cell yellow
            ActiveCell.Interior.Color = RGB(255, 255, 0)
        
        Else
            
            'Otherwise color the cell white
            ActiveCell.Interior.Color = RGB(255, 255, 255)
            
        End If
        
    Next Cell
    
    Application.ScreenUpdating = True
    
End Sub

'Quote Function:
Function Quote(Text)
    
    'Format the data
    Quote = Chr(34) & Text & Chr(34)
    
End Function

'Find Appointments Subroutine:
Sub Find_Appointments(Recipient, Starting_Cell, Starting_Date, Ending_Date)
    
    'Declare Needed Variables
    Dim Outlook_Application As Outlook.Application
    Dim Outlook_Namespace As Outlook.Namespace
    Dim Outlook_Calendar_Items As Outlook.Items
    Dim Outlook_Check_Items As Outlook.Items
    Dim Outlook_Appointment As Outlook.AppointmentItem
    Dim Outlook_Folder As Outlook.MAPIFolder
    Dim Calendar_Item As Outlook.AppointmentItem
    Dim Outlook_Recipient As Outlook.Recipient
    
    Dim Computer As String
    Dim Object_WMI_Service
    Dim Items_Collection
    Dim Object_Item
    
    Dim Item As Object
    Dim Check_String As String
    Dim Match As Object
    Dim Send_Text As String
    Dim Next_Row
    
    Dim Subject_Array() As String
    Dim Starting_Date_Array() As Variant
    Dim Starting_Time_Array() As Variant
    Dim Ending_Time_Array() As Variant
    Dim Text_Array() As String
    Dim Phone_Array() As Variant
    Dim Email_Array() As String
    Dim Advisor_Array() As String
    Dim Appointment_Type_Array() As String
    Dim Array_Size As Integer
    Dim i As Integer

    On Error Resume Next
        
    'Define the next row
    Next_Row = Cells(Rows.Count, 5).End(xlUp).Row
    
    'Define system variables
    Computer = "."
    Set Object_WMI_Service = GetObject("winmgmts:\\" & Computer & "\root\CIMV2")
    Set Items_Collection = Object_WMI_Service.ExecQuery("SELECT * FROM Win32_Process", , 48)
    
    'Search through the running programs
    For Each Object_Item In Items_Collection
    
        'Check for outlook
        If Object_Item.Name = "OUTLOOK.EXE" Then
            
            'If found open it
            Set Outlook_Application = GetObject(, "Outlook.Application")
            Exit For
            
        End If
        
    Next
    
    'Check if outlook was found
    If Outlook_Application Is Nothing Then
    
        'If not, open an outlook application
        Set Outlook_Application = CreateObject("Outlook.Application")
        
    End If
        
    'Check if outlook is still missing
    If Outlook_Application Is Nothing Then
    
        'If so, output an error message
        MsgBox "The email program has failed to start outlook. Please wait a moment and try again.", vbCritical, "Failed Outlook Startup"
        Exit Sub
            
    End If
        
    Application.Wait (Now + Time("00:00:05"))
    
    'Define outlook namespace
    Set Outlook_Namespace = Outlook_Application.GetNamespace("MAPI")
        
    'Define outlook recipient
    Set Outlook_Recipient = Outlook_Namespace.CreateRecipient(Recipient)
    Outlook_Recipient.Resolve
        
    On Error Resume Next
        
    'Define outlook calendar items
    Set Outlook_Calendar_Items = Outlook_Namespace.GetSharedDefaultFolder(Outlook_Recipient, olFolderCalendar).Items
        
    'Set the parameters for the calendar items
    With Outlook_Calendar_Items
        
        .Sort "[Start]", False
        .IncludeRecurrences = True
            
    End With
        
    'Enter the string to be checked
    Check_String = "[Start] >= " & Quote(Starting_Date & " 12:00 AM") & " AND [End] <= " & Quote(Ending_Date & " 11:59 PM")
        
    'Define check items
    Set Outlook_Check_Items = Outlook_Calendar_Items.Restrict(Check_String)
        
    'Check if there were any appointments
    If Outlook_Check_Items.Count > 0 Then
            
        'Define the Array_Size
        Array_Size = 1
        
        'Loop through each outlook item
        For Each Item In Outlook_Check_Items
                
            'Set default text status
            Send_Text = "No"
                
            'Check if the item is aan outlook appointment
            If Item.Class = olAppointment Then
                    
                'If so, set outlook appointment to equal the item
                Set Outlook_Appointment = Item
                
                'Check if the body of the appointment has a reason
                If InStr(1, Outlook_Appointment.body, "Reason:") >= 1 Then
                        
                    'If so, Check the secretary form settings
                    If Outlook_Appointment.GetInspector.ModifiedFormPages("Secretary").Controls("f").Value = "" Then
                            
                        Set Match = Regex_Pattern.Execute(Outlook_Appointment.body)
                            
                        If Not Match(0) = "" Then
                                
                            Outlook_Appointment.GetInspector.ModifiedFormPages("Secretary").Controls("f").Value = Match(0)
                            
                        Else
                            
                            Outlook_Appointment.GetInspector.ModifiedFormPages("Secretary").Controls("f").Value = "No Email Provided"
                            
                        End If
                        
                    End If
                        
                    'Check if the appointment wants a text reminder
                    If InStr(1, CStr(Outlook_Appointment.Categories), "Personal") >= 1 Or InStr(1, CStr(Outlook_Appointment.Categories), "Text Reminder") >= 1 Or Outlook_Appointment.GetInspector.ModifiedFormPages("Secretary").Controls("textReminder").Value Then
                            
                        'If so, set send text to yes and color the text cell yellow
                        Send_Text = "Yes"
                            
                    End If
                        
                    'Resize and fill each array
                    ReDim Preserve Subject_Array(Array_Size)
                    ReDim Preserve Starting_Date_Array(Array_Size)
                    ReDim Preserve Starting_Time_Array(Array_Size)
                    ReDim Preserve Ending_Time_Array(Array_Size)
                    ReDim Preserve Text_Array(Array_Size)
                    ReDim Preserve Phone_Array(Array_Size)
                    ReDim Preserve Email_Array(Array_Size)
                    ReDim Preserve Advisor_Array(Array_Size)
                    ReDim Preserve Appointment_Type_Array(Array_Size)
    
                    Subject_Array(Array_Size) = Outlook_Appointment.Subject
                    Starting_Date_Array(Array_Size) = Format(Outlook_Appointment.Start, "MM/DD/YYYY")
                    Starting_Time_Array(Array_Size) = Format(Outlook_Appointment.Start, "HH:MM AM/PM")
                    Ending_Time_Array(Array_Size) = Format(Outlook_Appointment.End, "HH:MM AM/PM")
                    Text_Array(Array_Size) = Send_Text
                    Phone_Array(Array_Size) = Outlook_Appointment.GetInspector.ModifiedFormPages("Secretary").Controls("e").Value + ","
                    Email_Array(Array_Size) = Outlook_Appointment.GetInspector.ModifiedFormPages("Secretary").Controls("f").Value
                    Advisor_Array(Array_Size) = Recipient
                    Appointment_Type_Array(Array_Size) = Outlook_Appointment.GetInspector.ModifiedFormPages("Secretary").Controls("1").Value
                     
                    Array_Size = Array_Size + 1
                End If
                    
            End If
            
        Next Item
            
    Else
        
        'If there are no appointments, output an error message
        MsgBox "There are no appoinments or meetings during the time you have specified.", vbInformation, "No Appointments"
            
    End If
    
    Outlook_Application.Close
    Set Outlook_Application = Nothing
    
    For i = 0 To Array_Size
    
        With Starting_Cell
        
            .End(xlDown).End(xlUp).Offset(Next_Row, 0).Value = Subject_Array(i)
            .End(xlDown).End(xlUp).Offset(Next_Row, 1).Value = Starting_Date_Array(i)
            .End(xlDown).End(xlUp).Offset(Next_Row, 2).Value = Starting_Time_Array(i)
            .End(xlDown).End(xlUp).Offset(Next_Row, 3).Value = Ending_Time_Array(i)
            .End(xlDown).End(xlUp).Offset(Next_Row, 4).Value = Text_Array(i)
            .End(xlDown).End(xlUp).Offset(Next_Row, 5).Value = Phone_Array(i)
            .End(xlDown).End(xlUp).Offset(Next_Row, 6).Value = Email_Array(i)
            .End(xlDown).End(xlUp).Offset(Next_Row, 7).Value = Advisor_Array(i)
            .End(xlDown).End(xlUp).Offset(Next_Row, 8).Value = Appointment_Type_Array(i)
            
        End With
        
        Next_Row = Next_Row + 1
        
    Next
        
    'Search through the running programs
    For Each Object_Item In Items_Collection
    
        'Check for outlook
        If Object_Item.Name = "OUTLOOK.EXE" Then
            
            'If found close it
            Object_Item.Terminate
            Exit For
                
        End If
        
    Next
        
End Sub

'Send Emails Subroutine:
Sub Send_Emails()

    'Declare needed variables
    Dim Workbook
    Dim Response
    Dim Row_Number As Long
    Dim Last_Row As Long
    Dim Appointment_Start_Time As String
    Dim Appointment_End_Time As String
    Dim Appointment_Day As String
    Dim Appointment_Type As String
    Dim Appointment_Location As String
    Dim Appointment_Information As String
    Dim Advisor As String
    Dim Email As String
    Dim Email_Message_Body As String
    Dim Subject_Line As String
    Dim Outlook_Application As Object
    Dim Outlook_Mail As Outlook.MailItem
    Dim Computer As String
    Dim Object_WMI_Service
    Dim Items_Collection
    Dim Object_Item
    
    'Get the correct workbook
    For Each Workbook In Workbooks
    
        'Check if the workbook is the new workbook
        If Workbook.Name = "Book*" Then
        
            'If so, activate the workbook
            Workbook.Activate
            Exit For
            
        End If
        
    Next Workbook

    'Get the active workbook
    Set Workbook = ActiveWorkbook
    
    'Check if there is any appointment data
    If Not Workbook.ActiveSheet.Range("A2").Value = "" Then
    
        'If so, ask the user if they'd like to continue
        Response = MsgBox("Would you like to send out the emails now? Please look over the data and make sure everything looks correct before proceeding.", vbInformation + vbYesNo, "Send Emails")
    
        'Check if the user wants to continue
        If Response = vbNo Then
        
            'If not, exit the sub and close the workbook
            MsgBox "Exiting the email program now.", vbInformation, "Exiting Program"
            Workbook.Close
            Exit Sub
        
        Else
        
            'Output a message informing the user that the emails are being sent
            MsgBox "The emails are now being sent. Please wait one moment.", vbInformation, "Sending Emails"
            
            Call Save_Workbook
            
            'Define the row number and last row
            Row_Number = 1
            Last_Row = Workbook.Worksheets("Sheet1").Range("A1").End(xlDown).Row
            
            'Loop through each email
            Do
            DoEvents
            
                'Define the subject line
                Subject_Line = "You have an appointment with Life Sciences Student Services"
                
                'Increment the row number
                Row_Number = Row_Number + 1
                
                'Get the appointment type
                Appointment_Type = Workbook.Worksheets("Sheet1").Range("I" & Row_Number)
                
                'Get and format appointment information
                Advisor = Workbook.Worksheets("Sheet1").Range("H" & Row_Number)
                Email = Workbook.Worksheets("Sheet1").Range("G" & Row_Number)
                Appointment_Day = Workbook.Worksheets("Sheet1").Range("B" & Row_Number)
                Appointment_Day = Format(Appointment_Day, "dddd, mmm d yyyy")
                Appointment_Start_Time = Workbook.Worksheets("Sheet1").Range("C" & Row_Number)
                Appointment_Start_Time = Format(Appointment_Start_Time, "hh:mm AMPM")
                Appointment_End_Time = Workbook.Worksheets("Sheet1").Range("D" & Row_Number)
                Appointment_End_Time = Format(Appointment_End_Time, "hh:mm AMPM")
                
                'Get the advisor's room number
                Select Case Advisor
                
                    Case "XXXX"
                        
                        Appointment_Location = "2060F LSB"
                        
                    Case "XXXX"
                    
                        Appointment_Location = "2060D LSB"
                        
                    Case "XXXX"
                    
                        Appointment_Location = "2060A LSB"
                        
                    Case "XXXX"
                    
                        Appointment_Location = "2060E LSB"
                        
                    Case "XXXX"
                    
                        Appointment_Location = "2060C LSB"
                        
                    Case Else
                        
                        Appointment_Location = "2060 LSB"
                
                End Select
                
                'Get the messages specific to different appointment types
                Select Case Appointment_Type
                
                    Case "none"
                        
                        Appointment_Information = "Additional Information:"
                    
                    Case "Academic Improvement Plan"
                    
                        Appointment_Information = "<h4>Academic Improvement Plan</h4> <p>Prior to coming to this appointment you need to have completed the Academic Improvement Plan (AIP), <a href=" & Chr(34) & _
                        "http://aip.byu.edu" & Chr(34) & ">Click here </a>to complete the AIP survey.</p>"
                        
                    Case "Major Exploration"
                    
                        Appointment_Information = "<h4>Major Exploration</h4> <p>In a major exploration appointment you will be able to explore the majors of our college with an Advisor, however; you cannot change your major in this appointment. If you are ready and would like to change your major you need to call 8014223042 and schedule a " & Chr(34) & "Change of Major" & Chr(34) & " appointment.</p>"
                        
                    Case "Change of Major"
                    
                        Appointment_Information = "<h4>Changing your Major:</h4><p>You will need to fill out a " & Chr(34) & "Change of Major" & Chr(34) & " request on MyMap. It can be found on the bottom left on the plan tab. Please complete this request prior to coming for your appointment.</p>"
                        
                    Case "Initial Class Plan"
                    
                        Appointment_Information = "Additional Information:"
                        
                    Case "Revise Class Plan"
                    
                        Appointment_Information = "Additional Information:"
                        
                    Case "Graduation Clearance"
                        
                        Subject_Line = "Important Convocation Information"
                        Appointment_Information = "<h4>If you want to walk for graduation, you must complete the following steps:</h4><br><p>1) Apply online for graduation:<a href=" & Chr(34) & "http://registrar.byu.edu/registrar/graduation/apply.php" & Chr(34) & _
                        "> Click Here to Apply for Graduation</a><br>2) If you plan on walking in a Life Sciences Convocation you must inform your advisor during this appointment in order to be added to the graduation walking database. If you fail to do this, there will not be a seat planned for you.</p><p>Please watch this video before meeting with your advisor:<a href=" & Chr(34) & _
                        "https://www.youtube.com/watch?v=DB6EZq22bhl" & Chr(34) & "> Click to view</a></p><p>Note: If you have a relative that would like to walk with you in the College of Life Sciences you will need to let your advisor know. This is how we will be able to seat the two of you together. Your relative's seating card will be clipped to yours when you arrive at convocation. Please send your advisor their first and last name, Net ID, student ID number, major, and e-mail so that we can save a seat for them and send them updates on graduation for the Life Sciences College.</p>"
                    
                    Case "Internship"
                    
                        Appointment_Information = "Additional Information:"
                        
                    Case "GE Questions"
                    
                        Appointment_Information = "Additional Information:"
                        
                    Case "GE Substitution"
                    
                        Appointment_Information = "Additional Information:"
                        
                    Case Else
                    
                        Appointment_Information = "Additional Information:"
                        
                End Select
                
                'Define the main email message
                Email_Message_Body = "<html><body><h3>Dear Student,</h3><p>This is a reminder that you have an appointment with an advisor from the Life Sciences Advisement Center:</p><ul><li>Date: " & Appointment_Day & _
                "</li><li>Time: " & Appointment_Start_Time & " - " & Appointment_End_Time & "</li><li>Reason: " & Appointment_Type & "</li><li>Advisor: " & Advisor & "</li><li>Where: " & Appointment_Location & "(Please disregard this if you're scheduled for a phone appointment)" & _
                "</li></ul></p>" & Appointment_Information & "<p>Thank you for coming prepared to your appointment! If you cannot make your appointment time you may respond to this email and cancel or call (801)422-3042 to reschedule. If you are meeting in-office with an advisor and come more than 5 minutes late you may be asked to reschedule. If you have a phone appointment your advisor will call you at the appointed time. If your advisor has not called you within 5-10 minutes of your scheduled time, please call the front desk at (801) 422-3042</p><p>As a reminder, we expect that you will comply with BYU's Dress and Grooming Standards for your appointment. If you fail to meet these standards, you will be asked to reschedule. Thank you!</p><p>Life Sciences Student Services<br>2060 LSB<br>lifesciences@byu.edu<br><a href=" & _
                Chr(34) & "https://www.facebook.com/lifesciences.byu.edu" & Chr(34) & "><img src=" & Chr(34) & "https://cdn-images.mailchimp.com/icons/social-block/color-facebook-128.png" & Chr(34) & " alt=" & Chr(34) & "Facebook" & Chr(34) & " height=" & Chr(34) & "42" & Chr(34) & _
                " width=" & Chr(34) & "42" & Chr(34) & "></a><a href=" & Chr(34) & "https://twitter.com/lifesciencesbyu" & Chr(34) & "><img src=" & Chr(34) & "https://cdn-images.mailchimp.com/icons/social-block/color-twitter-128.png" & Chr(34) & " alt=" & Chr(34) & _
                "Twitter" & Chr(34) & " height=" & Chr(34) & "42" & Chr(34) & " width=" & Chr(34) & "42" & Chr(34) & "></a><a href=" & Chr(34) & "mailto:lifesciences@byu.edu" & Chr(34) & "><img src=" & Chr(34) & "https://cdn-images.mailchimp.com/icons/social-block/color-forwardtofriend-128.png" & Chr(34) & _
                " alt=" & Chr(34) & "Email" & Chr(34) & " height=" & Chr(34) & "42" & Chr(34) & " width=" & Chr(34) & "42" & Chr(34) & "></a></p></body></html>"
                
                'Check if the the student has an email listed
                If Not Email = "" Then
                
                    If Row_Number = 2 Then
                        
                        'Define system variables
                        Computer = "."
                        Set Object_WMI_Service = GetObject("winmgmts:\\" & Computer & "\root\CIMV2")
                        Set Items_Collection = Object_WMI_Service.ExecQuery("SELECT * FROM Win32_Process", , 48)
                        
                        'Search through the running programs
                        For Each Object_Item In Items_Collection
                        
                            'Check for outlook
                            If Object_Item.Name = "OUTLOOK.EXE" Then
                                
                                'If found open it
                                On Error Resume Next
                                Set Outlook_Application = GetObject(, "Outlook.Application")
                                Exit For
                                
                            End If
                            
                        Next
                        
                        'Check if outlook was found
                        If Outlook_Application Is Nothing Or Err Then
                        
                            'If not, open an outlook application
                            Err.Clear
                            Set Outlook_Application = CreateObject("Outlook.Application")
                            
                        End If
                            
                        'Check if outlook is still missing
                        If Outlook_Application Is Nothing Then
                        
                            'If so, output an error message
                            MsgBox "The email program has failed to start outlook. Please wait a moment and try again.", vbCritical, "Failed Outlook Startup"
                            Exit Sub
                                
                        End If
                    
                    End If
                    
                    'Create an email draft
                    Set Outlook_Mail = Outlook_Application.CreateItem(olMailItem)
                    
                    'Fill out the to line
                    Outlook_Mail.To = Email
                    
                    'Fill out the subject line
                    Outlook_Mail.Subject = Subject_Line
                    
                    'Fill in and format the body of the email
                    Outlook_Mail.body = Email_Message_Body
                    Outlook_Mail.BodyFormat = olFormatHTML
                    Outlook_Mail.HTMLBody = Email_Message_Body
                    
                    'Send the email
                    Outlook_Mail.send
                
                    'Mark that an email was sent
                    Workbook.ActiveSheet.Range("G" & Row_Number).Activate
                    ActiveCell.Interior.Color = RGB(66, 215, 90)
                
                End If
                
            Loop Until Row_Number = Last_Row
            
            'Ask the user if they'd like to close the workbook
            Response = MsgBox("The emails have been sent. Would you like to close this document now?", vbYesNo + vbInformation, "Updating Complete")
        
            'Check if the user said yes
            If Response = vbYes Then
            
                'If so, close the workbook
                Workbook.Save
                Workbook.Close
                
            End If
            
        End If
        
    Else
    
        'Output an error message, close the workbook, and exit the subroutine
        MsgBox "There are no appointments listed. If this is a mistake, please rerun the program and try again.", vbInformation, "No Appointments"
        Workbook.Close
        Exit Sub
    
    End If

End Sub

'Send Survey Subroutine:
Sub Send_Survey()

    'Declare needed variables
    Dim Workbook
    Dim Response
    Dim Row_Number As Long
    Dim Last_Row As Long
    Dim Email As String
    Dim Advisor As String
    Dim Email_Message_Body As String
    Dim Subject_Line As String
    Dim Outlook_Application As Outlook.Application
    Dim Outlook_Mail As Outlook.MailItem
    Dim Computer As String
    Dim Object_WMI_Service
    Dim Items_Collection
    Dim Object_Item
    Dim Save_Date As String
    Dim Workbook_Name As String
    Dim Path As String
    
    'Get the correct workbook
    For Each Workbook In Workbooks
    
        'Check if the workbook is the new workbook
        If Workbook.Name = "Book*" Then
        
            'If so, activate the workbook
            Workbook.Activate
            Exit For
            
        End If
        
    Next Workbook

    'Get the active workbook
    Set Workbook = ActiveWorkbook

    'Check if there is any appointment data
    If Not Workbook.ActiveSheet.Range("A2").Value = "" Then
    
        'If so, ask the user if they'd like to continue
        Response = MsgBox("Would you like to send out the emails now? Please look over the data and make sure everything looks correct before proceeding.", vbInformation + vbYesNo, "Send Emails")
    
        'Check if the user wants to continue
        If Response = vbNo Then
        
            'If not, exit the sub and close the workbook
            MsgBox "Exiting the email program now.", vbInformation, "Exiting Program"
            Workbook.Close
            Exit Sub
        
        Else
        
            'Output a message informing the user that the emails are being sent
            MsgBox "The emails are now being sent. Please wait one moment.", vbInformation, "Sending Emails"
            
            'Send the emails
            Call Save_Workbook
            
            'Define the row number and last row
            Row_Number = 1
            Last_Row = Workbook.Worksheets("Sheet1").Range("A1").End(xlDown).Row
            
            'Loop through each email
            Do
            DoEvents
            
                'Define the subject line
                Subject_Line = "Life Sciences Student Services Visit Survey"
                
                'Increment the row number
                Row_Number = Row_Number + 1
                
                'Get the student's email and the advisor they met with
                Email = Workbook.Worksheets("Sheet1").Range("G" & Row_Number)
                Advisor = Workbook.Worksheets("Sheet1").Range("H" & Row_Number)
                                
                'Define the main email message
                Email_Message_Body = "<html><body><h3>Dear Student, </h3><p>We understand that you have recently had an appointment scheduled with your advisor, " & Advisor & _
                ", at the Life Science Student Services Center 2060 LSB. If you met with your advisor, it would be greatly appreciated if you could take a few minutes and complete this survey to help us better assist you and the other students.</p><p>If you received this email by mistake we apologize for the inconvenience.</p><p>https://byu.az1.qualtrics.com/SE/?SID=SV_20q3KlKbU0wOdtH</p><p>Thank You,</p><p>Life Sciences Student Services<br>2060 LSB<br>lifesciences@byu.edu<br><a href=" & _
                Chr(34) & "https://www.facebook.com/lifesciences.byu.edu" & Chr(34) & "><img src=" & Chr(34) & "https://cdn-images.mailchimp.com/icons/social-block/color-facebook-128.png" & Chr(34) & " alt=" & Chr(34) & "Facebook" & Chr(34) & " height=" & Chr(34) & "42" & Chr(34) & _
                " width=" & Chr(34) & "42" & Chr(34) & "></a><a href=" & Chr(34) & "https://twitter.com/lifesciencesbyu" & Chr(34) & "><img src=" & Chr(34) & "https://cdn-images.mailchimp.com/icons/social-block/color-twitter-128.png" & Chr(34) & " alt=" & Chr(34) & _
                "Twitter" & Chr(34) & " height=" & Chr(34) & "42" & Chr(34) & " width=" & Chr(34) & "42" & Chr(34) & "></a><a href=" & Chr(34) & "mailto:lifesciences@byu.edu" & Chr(34) & "><img src=" & Chr(34) & "https://cdn-images.mailchimp.com/icons/social-block/color-forwardtofriend-128.png" & Chr(34) & _
                " alt=" & Chr(34) & "Email" & Chr(34) & " height=" & Chr(34) & "42" & Chr(34) & " width=" & Chr(34) & "42" & Chr(34) & "></a></p></body></html>"
                
                'Check if the the student has an email listed
                If Not Email = "" Then
                
                    If Row_Number = 2 Then
                    
                        'Define system variables
                        Computer = "."
                        Set Object_WMI_Service = GetObject("winmgmts:\\" & Computer & "\root\CIMV2")
                        Set Items_Collection = Object_WMI_Service.ExecQuery("SELECT * FROM Win32_Process", , 48)
                        
                        'Search through the running programs
                        For Each Object_Item In Items_Collection
                        
                            'Check for outlook
                            If Object_Item.Name = "OUTLOOK.EXE" Then
                                
                                'If found open it
                                On Error Resume Next
                                Set Outlook_Application = GetObject(, "Outlook.Application")
                                Exit For
                                
                            End If
                            
                        Next
                        
                        'Check if outlook was found
                        If Outlook_Application Is Nothing Or Err Then
                        
                            'If not, open an outlook application
                            Err.Clear
                            Set Outlook_Application = CreateObject("Outlook.Application")
                            
                        End If
                            
                        'Check if outlook is still missing
                        If Outlook_Application Is Nothing Then
                        
                            'If so, output an error message
                            MsgBox "The email program has failed to start outlook. Please wait a moment and try again.", vbCritical, "Failed Outlook Startup"
                            Exit Sub
                                
                        End If
                        
                    End If
        
                    'Create an email draft
                    Set Outlook_Mail = Outlook_Application.CreateItem(olMailItem)
                    
                    'Fill out the to line
                    Outlook_Mail.To = Email
                    
                    'Fill out the subject line
                    Outlook_Mail.Subject = Subject_Line
                    
                    'Fill in and format the body of the email
                    Outlook_Mail.body = Email_Message_Body
                    Outlook_Mail.BodyFormat = olFormatHTML
                    Outlook_Mail.HTMLBody = Email_Message_Body
                    
                    'Send the email
                    Outlook_Mail.send
                    
                    'Mark that an email was sent
                    Workbook.ActiveSheet.Range("G" & Row_Number).Activate
                    ActiveCell.Interior.Color = RGB(66, 215, 90)
                        
                    
                End If
               
            Loop Until Row_Number = Last_Row
            
        End If
        
        'Output a message informing the user that the workbook is being saved
        MsgBox "The surveys have been sent. Wait a moment while this workbook is saved.", vbInformation, "Saving Workbook"
        
        'Define the needed information to save the workbook
        Save_Date = Format(DateTime.Now, "MM-DD-YYYY HH.MM AM/PM")
        Workbook_Name = "Appointments_" & Save_Date
        Path = "M:\Secretary Resources\AppTracking\" & Workbook_Name
        
        'Save the workbook
        ActiveWorkbook.SaveAs Path
        
        'Ask if the user wants to close the workbook
        Response = MsgBox("The workbook has been saved. Would you like to close this workbook now?", vbYesNo + vbInformation, "Save Complete")
        
        'Check if they said yes
        If Response = vbYes Then
        
            'If so, close the workbook
            Workbook.Close
            
        End If
        
    Else
    
        'Output an error message, close the workbook, and exit the subroutine
        MsgBox "There are no appointments listed. If this is a mistake, please rerun the program and try again.", vbInformation, "No Appointments"
        Workbook.Close
        Exit Sub
    
    End If

End Sub

'Save Workbook Subroutine:
Sub Save_Workbook()

    'Declare needed variables
    Dim Save_Date
    Dim Save_Path

    'Format the date
    Save_Date = Format(DateTime.Now, "MM-DD-YYYY HH.MM AM/PM")
    
    'Define the path to save the file
    Save_Path = "M:\Secretary Resources\~ReminderEmails\SentEmails\SentEmails_" & Save_Date & ".xlsx"
    
    'Save the file
     ActiveWorkbook.SaveAs Save_Path
    
End Sub

