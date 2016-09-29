'========================= BYU SIGN-IN/LOGOUT ========================='
'Created By: Brady Hammond
'Created: 05/04/16
'======================================================================'

'BYU Login Subroutine:
Sub BYU_LOGIN()

    'Declare needed variables
    Dim HTML_Document As HTMLDocument
    Dim IE_Browser As Object
    Dim BYU_URL As String
    Dim URL_Check As String
    Dim IE_Check As SHDocVw.InternetExplorer
    Dim IE_Check_Shell_Window As SHDocVw.ShellWindows
    
    'Create a shell window
    Set IE_Check_Shell_Window = New SHDocVw.ShellWindows
    
    'Check if there are open windows
    If IE_Check_Shell_Window.Count Then
        
        'Loop through each open window
        For Each IE_Check_Object In IE_Check_Shell_Window
                       
            'Close and delete each one
            IE_Check_Object.Quit
            Set IE_Check_Object = Nothing
            
        Next IE_Check_Object
        
    End If
    
    'Define BYU_URL and URL_Check
    BYU_URL = "https://my.byu.edu/uPortal/f/u20l1s5/normal/render.uP"
    
    'Open up internet explorer
    Set IE_Browser = CreateObject("InternetExplorer.Application")
    
    On Error GoTo Error_Handling
    
    'Disable pop-ups
    IE_Browser.Silent = True
    
    'Navigate to the BYU login page
    IE_Browser.Navigate BYU_URL
    
    'Hide the browser
    IE_Browser.Visible = False
    
    'Wait for the page to load
    Do
    Loop Until IE_Browser.readyState = 3
    
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Check if the user is already logged into a myBYU account
    If IE_Browser.LocationURL = BYU_URL Then
    
        'If so, exit the browser
        IE_Browser.Quit
        
        'Delete the browser object
        Set IE_Browser = Nothing
        
        'Output that the user is already logged off
        MsgBox "You are already logged into your myBYU account.", vbInformation, "Already Signed-in"
        
        'Exit the subroutine
        Exit Sub
        
    Else
    
        'Show the BYU login form
        BYU_Login_Form.Show
    
    End If
    
Error_Handling:
    
    If Err.Number <> 0 Then
        
        MsgBox "The scanning program has failed to connect to the internet. Please wait a moment and try again.", vbCritical, "Failed Internet Connection"
        
    End If
    
End Sub
Sub BYU_LOGOUT()

    'Declare needed variables
    Dim HTML_Document As HTMLDocument
    Dim IE_Browser As Object
    Dim HTML_Element As IHTMLElement
    Dim BYU_URL As String
    Dim URL_Check As String
    Dim Extra_Windows_Check_Shell_Window As SHDocVw.ShellWindows
    Dim Extra_Windows_Check_Object As Object

    'Create a shell window
    Set Extra_Windows_Check_Shell_Window = New SHDocVw.ShellWindows
    
    'Define BYU_URL and URL_Check
    BYU_URL = "https://my.byu.edu/uPortal/f/u20l1s5/normal/render.uP"
    URL_Check = "https://my.byu.edu/uPortal/f/u26l1s4/normal/render.uP"
    
    'Open up internet explorer
    Set IE_Browser = CreateObject("InternetExplorer.Application")
    
    On Error GoTo Error_Handling
    
    'Disable pop-ups
    IE_Browser.Silent = True
    
    'Navigate to the BYU login page
    IE_Browser.Navigate BYU_URL
    
    'Hide the browser
    IE_Browser.Visible = False
    
    'Wait for the page to load
    Do
    Loop Until IE_Browser.readyState = 3
    
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Check if the user is already logged into a myBYU account
    If IE_Browser.LocationURL <> BYU_URL Then
    
        'If so, exit the browser
        IE_Browser.Quit
        
        'Delete the browser object
        Set IE_Browser = Nothing
        
        'Output that the user is already logged off
        MsgBox "You are already logged out of your myBYU account.", vbInformation, "Already Signed-out"
        
        'Exit the subroutine
        Exit Sub
        
    Else
    
        'Save the HTML
        Set HTML_Document = IE_Browser.document
        
        'Loop through each HTML element
        For Each HTML_Element In HTML_Document.getElementsByTagName("span")
        
            'Check if the sign out button is found
            If HTML_Element.innerHTML = "Sign Out" Then
            
                'If it is, click it and exit the loop
                HTML_Element.Click: Exit For
                
            End If
            
        Next HTML_Element
            
        'Wait for the page to load
        Do
        Loop Until IE_Browser.readyState = 3
        
        Do
        Loop Until IE_Browser.readyState = 4
        
        'Check if the logout worked
        If IE_Browser.LocationURL = URL_Check Then
            
                'If it did, exit the browser
                IE_Browser.Quit
            
                'Delete the browser object
                Set IE_Browser = Nothing
            
                'Output a success message
                MsgBox "You have successfully signed out of your myBYU account.", vbInformation, "Successful Sign-out"
            
        Else
        
            'Exit the browser
            IE_Browser.Quit
            
            'Delete the browser object
            Set IE_Browser = Nothing
            
            'Output an error message
            MsgBox "Something went wrong. Forcing logout of your myBYU account.", vbCritical, "Sign-in Error"
            
        End If
        
    End If
    
    'Check if there are open windows
    If Extra_Windows_Check_Shell_Window.Count Then
    
        'Loop through each open window
        For Each Extra_Windows_Check_Object In Extra_Windows_Check_Shell_Window
        
            'Exit the pop up browser
            Extra_Windows_Check_Object.Quit
                        
            'Delete the pop up
            Set Extra_Windows_Check_Object = Nothing
            
        Next Extra_Windows_Check_Object
            
        On Error Resume Next
        Err.Clear
        
    End If
    
Error_Handling:
    
    If Err.Number <> 0 Then
        
        MsgBox "The scanning program has failed to connect to the internet. Please wait a moment and try again.", vbCritical, "Failed Internet Connection"
        
    End If
    
End Sub
