'========================= LOGIN FORM ========================='
'Created By: Brady Hammond
'Created: 05/04/16
'=============================================================='

'Initialize User Form Subroutine:
Private Sub UserForm_Initialize()

    'Empty NetID box
    NetID_Input.Value = ""
    
    'Empty Password box
    Password_Input.Value = ""
    
End Sub

'Submit Button (Enter) Subroutine:
Private Sub Submit_Button_Enter()
    
    'Set the focus on the submit button and click it
    Submit_Button.SetFocus
    Submit_Button.Value = True
    
End Sub

'Sign-In Button Subroutine:
Private Sub Submit_Button_Click()
    
    'Declare needed variables
    Dim NetID As String
    Dim Password As String
    Dim HTML_Document As HTMLDocument
    Dim IE_Browser As Object
    Dim HTML_Element As IHTMLElement
    Dim BYU_URL As String
    Dim URL_Check As String
    
    'Save NetID and Password
    NetID = NetID_Input.Value
    Password = Password_Input.Value
    
    'Define BYU_URL and URL_Check
    BYU_URL = "https://cas.byu.edu/cas/login?service=https://my.byu.edu/uPortal/Login"
    URL_Check = "https://cas.byu.edu/cas/login?service=https://my.byu.edu/uPortal/Login"
    
    
    'Check if the user has input a value for NetID
    If NetID = "" Then
    
        'If not, output an error message and clear the form
        MsgBox ("Please enter your NetID and Password")
        Call UserForm_Initialize
    
    'Check if the user has input a value for Password
    ElseIf Password = "" Then
    
        'If notm output an error message and clear the form
        MsgBox ("Please enter your NetID and Password")
        Call UserForm_Initialize
           
    'Continue if the user has entered both fields
    Else
    
        'Open up internet explorer
        Set IE_Browser = CreateObject("InternetExplorer.Application")
        
        'Disable pop-ups
        IE_Browser.Silent = True
        
        'Navigate to the BYU login page
        IE_Browser.Navigate BYU_URL
        
        'Hide the browser
        IE_Browser.Visible = False
        
        'Wait for the page to load
        Do
        Loop Until IE_Browser.readyState = 4
        
        'Save the HTML
        Set HTML_Document = IE_Browser.document
        
        'Find netid in the HTML and input NetID
        HTML_Document.getElementById("netid").Value = NetID
        
        'Find password in the HTML and input Password
        HTML_Document.getElementById("password").Value = Password
        
        'Loop through each HTML element
        For Each HTML_Element In HTML_Document.getElementsByTagName("input")
            
            'Check if the submit button is found
            If HTML_Element.Type = "submit" Then
            
                'If it is, click it and exit the loop
                HTML_Element.Click: Exit For
                
            End If
            
        Next HTML_Element
                
        'Wait for the page to load
        Do While IE_Browser.Busy
        Loop
        
        'Continue despite errors
        On Error Resume Next
        
        'Check if the login worked
        If IE_Browser.LocationURL = URL_Check Then
        
            'If it didn't, exit the browser
            IE_Browser.Quit
            
            'Delete the browser object
            Set IE_Browser = Nothing
            
            'Output an error message
            MsgBox "Then NetID or Password you entered was not valid. Please try again.", vbCritical, "Incorrect Login"
            
            'Clear the NetID
            NetID = ""
            
            'Clear the Password
            Password = ""
            
            'Clear the input boxes
            Call UserForm_Initialize
            
            'Exit the subroutine
            Exit Sub
                
        Else
            
            'Output a success message
            MsgBox "You are now logged into your myBYU account.", vbInformation, "Successful Login"
            
            'Unload the user form
            Unload Me
            
        End If
        
    End If
    
End Sub

'Cancel Button Subroutine:
Private Sub Cancel_Button_Click()
    
    'Unload the user form
    Unload Me
    
End Sub
