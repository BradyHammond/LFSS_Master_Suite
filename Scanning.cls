'========================= SCANNING FORM ========================='
'Created By: Brady Hammond
'Created: 05/05/16
'================================================================='

'Scanning Subroutine
Sub SCANNING_PROGRAM()

 'Declare needed variables
    Dim IE_Browser As Object
    Dim HTML_Document As HTMLDocument
    Dim HTML_Element As IHTMLElement
    Dim BYU_URL As String
    Dim URL_Check As String
    Dim Secretary As String
    Dim Files_To_Be_Scanned As String
    Dim Student_Digital_Copies As String
    Dim Last_File_To_Delete As String
    Dim Extra_Windows_Check_Shell_Window As SHDocVw.ShellWindows
    Dim Extra_Windows_Check_Object As Object
    
    'Create a shell window
    Set Extra_Windows_Check_Shell_Window = New SHDocVw.ShellWindows
    
    'Check if there are open windows
    If Extra_Windows_Check_Shell_Window.Count Then
    
        'Loop through each open window
        For Each Extra_Windows_Check_Object In Extra_Windows_Check_Shell_Window
        
            If InStr(1, Extra_Windows_Check_Object.LocationName, "Brigham Young University Sign-in Service") Then
        
                'Exit the pop up browser
                Extra_Windows_Check_Object.Quit
                        
                'Delete the pop up
                Set Extra_Windows_Check_Object = Nothing
            
            End If
            
        Next Extra_Windows_Check_Object
            
        On Error Resume Next
        Err.Clear
        
    End If
    
    'Define BYU_URL and URL_Check
    BYU_URL = "https://y.byu.edu/ry/ae/prod/person/cgi/personSummary.cgi"
    URL_Check = "https://cas.byu.edu/cas/login?service=https%3A%2F%2Fy.byu.edu%2Fvalidate%3Ftarget%3Dhttps%253A%252F%252Fy.byu.edu%252Fry%252Fae%252Fprod%252Fperson%252Fcgi%252FpersonSummary.cgi"

    'Open up internet explorer
    Set IE_Browser = CreateObject("InternetExplorer.Application")
    
    'Disable pop-ups
    IE_Browser.Silent = True
    
    'Navigate to the BYU login page
    IE_Browser.Navigate BYU_URL
    
    'Hide the browser
    IE_Browser.Visible = False
    
    'Wait for the page to load state 4
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Check if the user is already logged into a myBYU account
    If IE_Browser.LocationURL <> BYU_URL Or IE_Browser.LocationURL = URL_Check Then
    
        'If so, exit the browser
        IE_Browser.Quit
        
        'Delete the browser object
        Set IE_Browser = Nothing
        
        'Output that the user is already logged off
        MsgBox "You are not logged into your myBYU account. Please login and try again.", vbInformation, "Invalid User"
        
        'Exit the subroutine
        Exit Sub
        
    Else
    
        'Exit and delete the browser
        IE_Browser.Quit
        Set IE_Browser = Nothing
                
        'Run the scanning program
        Call Run_Scanning(Secretary, Files_To_Be_Scanned, Student_Digital_Copies, Last_File_To_Delete)
        
        'Get the correct workbook
        For Each Workbook In Workbooks
    
            'Check if the workbook is the created workbook
            If Workbook.Name = "Book*" Then
        
                'If so, activate the workbook
                Workbook.Activate
                
                'Autofit the data
                ActiveWorkbook.ActiveSheet.Columns("A:H").AutoFit
                
                'Turn on screen updating
                Application.ScreenUpdating = True
            
                'Delete the files in to be scanned
                Call Delete_Files("M:\To_Be_Scanned\")
                
                'Close the excel sheet
                Call Close_Program
                
                Exit For
            
            End If
        
        Next Workbook
        
        Err.Clear
        
    End If
    
    Err.Clear
    
Error_Handling:

    'Check if there are any errors
    If Err.Number <> 0 Then
        
        'If so, output an error message
        MsgBox "The scanning program has lost its connection to the internet. The program will now close. Please wait a moment and try again.", vbCritical, "Failed Internet Connection"
        Unload Scanning_User_Form
        
    End If
    
End Sub

Private Sub Run_Scanning(Secretary, Files_To_Be_Scanned, Student_Digital_Copies, Last_File_To_Delete)
    
    'Declare needed variables
    Dim Cell
    Dim File
    Dim Found_Files
    Dim Number_Of_Files As Integer
    
    'Clear input variables
    Last_File_To_Delete = ""
    Student_Digital_Copies = ""
    Files_To_Be_Scanned = ""
    
    'Check if the M Drive exists
    If Not Drive_Exists Then
    
        'If it does not, output an error message and exit the subroutine
        MsgBox "The M-Drive is disconnected. Please contact IT or try again later.", vbCritical, "Failed to Locate Drive"
        Exit Sub
   
    'Check if base folders set correctly
    ElseIf Not Set_Base_Folders(Files_To_Be_Scanned, Student_Digital_Copies) Then
        
        'Exit the subroutine
        Exit Sub
        
    Else
        
        'Open a new workbook
        Dim Workbook
        Set Workbook = Workbooks.Add
        
        With Workbook
            
            'Name the workbook
            .Title = "Scanning Workbook"
            
            'Create a subject
            .Subject = "Scanning Form"
            
            'Activate the workbook
            .Activate
            
            'Hide the forumla bar and headings
            .Application.DisplayFormulaBar = False
            ActiveWindow.DisplayHeadings = False
            
            With .Sheets(1).Range("A1:H1")
            
                'Format the workbook
                .HorizontalAlignment = xlCenter
                .ColumnWidth = 15
                .Interior.ColorIndex = 42
                .Font.Name = "Trebuchet MS"
                .Font.Size = 12
                .Font.ColorIndex = 2
                .Borders(xlBottom).Weight = xlThick
                .Borders.ColorIndex = 31
                
            End With
            
            With .Sheets(1).Range("A1")
                
                'Set up the headers for the workbook
                .Value = "Filename"
                .Offset(0, 1).Value = "New Filename"
                .Offset(0, 2).Value = "ID number"
                .Offset(0, 3).Value = "Student Name"
                .Offset(0, 4).Value = "Status"
                .Offset(0, 5).Value = "Datestamp"
                .Offset(0, 6).Value = "Secretary"
                .Offset(0, 7).Value = "Location"
                
            End With
        
        End With
        
        'Output an instruction box
        MsgBox "Please select the folder of files that you would like to scan.", vbInformation + vbOKOnly, "Folder Selection"

Folder_Selection:

        'Check if the to be scanned folder exists
        If Report_Folder_Status("M:\To_Be_Scanned\") Then
            
            'If so, set files to be scanned to that folder
            Files_To_Be_Scanned = Get_Folder("M:\To_Be_Scanned\")
            
        'Check if the m drive exists
        ElseIf Report_Folder_Status("M:\") Then
            
            'If so, set files to be scanned to that folder
            Files_To_Be_Scanned = Get_Folder("M:\")
            
        Else
            
            'Set files to be scanned to the main drive
            Files_To_Be_Scanned = Get_Folder("C:\")
            
        End If
        
        'Check if the folder has a slash at the end
        If Right(Files_To_Be_Scanned, 1) <> "\" Then
        
            'If not, add one
            Files_To_Be_Scanned = Files_To_Be_Scanned & "\"
            
        End If
        
        'Check if the selected folder exists
        If Report_Folder_Status(Files_To_Be_Scanned) Then
        
            'Upload all the pdfs from the selected folder
            Found_Files = Get_Files(Files_To_Be_Scanned)
            
        Else
        
            'Output an error message
            MsgBox "Invalid Selection please try again.", vbCritical + vbOKOnly, "Selection Error"
            
            GoTo Folder_Selection
            
        End If
        
        'Check if any files were returned
        If Found_Files = 0 Then
        
            'If not, output an error message, close the workbook, and exit the subroutine
            MsgBox "No PDF files to be scanned. Exiting scanning procedure now.", vbInformation, "No files found"
            ActiveWorkbook.Close SaveChanges:=False
            Exit Sub
            
        End If
        
        'Check if files were saved to the workbook
        If Workbook.Sheets(1).Range("A2").Value = "" Then
            
            'If not, output an error message, close the workbook, and exit the subroutine
            MsgBox "It seems that no files were found in M:\To_Be_Scanned. Exiting scanning procedure now.", vbInformation, "No files found"
            ActiveWorkbook.Close SaveChanges:=False
            Exit Sub
            
        End If
        
        'Find the number of files
        Number_Of_Files = Workbook.Sheets(1).Range("A1").End(xlDown).Row
        
        'Set cell to 2
        Cell = 2
        
        'Save the value in the first cell
        File = Workbook.Sheets(1).Range("A" & Cell).Value
        
        'Check if there is a value in the first cell
        If File = "" Then
        
            'If not, output an error message, close the workbook, and exit the subroutine
            MsgBox "No PDFs to File.", vbInformation + vbOKOnly, "Scanning Form"
            ActiveWorkbook.Close SaveChanges:=False
            Exit Sub
            
        Else
        
            'Get the secretary name and show the user form
            Secretary = Get_Secretary(Secretary)
            Call Show_Scanning_User_Form(File, Cell, Number_Of_Files, Files_To_Be_Scanned, Student_Digital_Copies)
            
            'Delete the browser
            IE_Browser.Quit
            Set IE_Browser = Nothing
            
            Workbooks("Book*.xls").Activate
            
        End If
        
    End If
    
End Sub

'Drive Exist Function:
Function Drive_Exists()

    'Declare and define the variables needed for this function
    Dim File_System_Object
    Dim Drive
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    Drive = "M:\"
    
    'Check if the file exists
    If File_System_Object.DriveExists(Drive) Then

        'If it  does, return true
        Drive_Exists = True
        
    Else
    
        'If it does not, return false
        Drive_Exists = False
        
    End If
    
End Function

'Set Base Folders Function:
Function Set_Base_Folders(Files_To_Be_Scanned, Student_Digital_Copies)

    'Declare needed variables
    Dim Scanning_Folder
    Dim Student_Folders
    
    'Define needed folders
    Scanning_Folder = "M:\To_Be_Scanned\"
    Student_Folders = "M:\Student Digital Copies\"

    'Declare and define the variables needed for this function
    Set_Base_Folders = True
    
    'Check that the folder exists
    If Report_Folder_Status(Scanning_Folder) Then
        
        'If it does, save it as toBeScanned
        Files_To_Be_Scanned = Scanning_Folder
        
    Else
    
        'If it doesn't, output an error message and return false
        MsgBox Scanning_Folder & " does not exist. Exiting now.", vbCritical, "Folder doesn't Exist"
        Set_Base_Folders = False
        
    End If
    
    'Check that the folder exists
    If Report_Folder_Status(Student_Folders) Then
    
        'If it does, save it as studentDigitalCopies
        Student_Digital_Copies = Student_Folders
        
    Else
    
        'If it does not, output an error message and set setBaseFolders to false
        MsgBox Student_Folders & " does not exist. Exiting now.", vbCritical, "Folder doesn't Exist"
        Set_Base_Folders = False
        
    End If
    
End Function

'Report File Status Function:
Function Report_Folder_Status(File)

    'Declare and define needed variables
    Dim File_System_Object
    Dim Status
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
   
    'Check if the file exists
    If (File_System_Object.FolderExists(File)) Then
      
        'If it does, set status to true
        Status = True
        
    Else
   
        'If it doesn't, set status to false
        Status = False
        
   End If
   
   'Return the file status
   Report_Folder_Status = Status
   
End Function

'Get Folder Function:
Function Get_Folder(Path As String) As String

    'Declare and define needed variables
    Dim Folder As FileDialog
    Dim Selected_Item As String
    Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
    
    With Folder
        
        'Add a title
        .Title = "Select a Folder"
        
        'Disable mulitselections
        .AllowMultiSelect = False
        
        'Set the initial file name
        .InitialFileName = Path
        
        'Check that show is not false
        If .Show <> -1 Then
            
            'If so, go to the NextCode
            GoTo Next_Code
        End If
        
        'Save the selected item
        Selected_Item = .SelectedItems(1)
    End With
    
    
Next_Code:
    
    'Return the selected item
    Get_Folder = Selected_Item

    'Delete the folder
    Set Folder = Nothing
    
End Function

'Get Files Function:
Function Get_Files(Files_To_Be_Scanned)

    'Declare needed variables
    Dim File_System_Object As Object
    Dim Scanning_Folder
    Dim Worksheet As Worksheet
    Dim Extention
    Dim File
    Dim Counter
    
    'Define needed variables
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    Set Scanning_Folder = File_System_Object.GetFolder(Files_To_Be_Scanned)
    Set Worksheet = ActiveWorkbook.ActiveSheet
    Extention = ""
    Counter = 2
    
    'Loop through each file
    For Each File In Scanning_Folder.Files
    
        'Save the file extension
        Extention = Right$(File.Path, Len(File.Path) - InStrRev(File.Path, "."))
        
        'Check for pdf files
        If Extention = "pdf" Then
        
            'If there is one, save it on the worksheet
            Worksheet.Range("A" & Counter).Value = File.Name
            
            'Increment the counter
            Counter = Counter + 1
            
        End If
        
    Next
    
    'Delete File_System_Object, Scanning_Folder, and Worksheet
    Set File_System_Object = Nothing
    Set Scanning_Folder = Nothing
    Set Worksheet = Nothing
    
    'Return the counter - 2
    Get_Files = Counter - 2
    
End Function

'Get Secretary Function:
Function Get_Secretary(Secretary) As String
    
    Do
        'Check if there is a value for secretary
        If Secretary = "" Then
        
            'If not, open an input box
            Secretary = InputBox("Please enter your First and Last name.", "Secretary for Scanning Batch")
            
            'Save input as Secretary
            Scanning_User_Form.Secretary.Caption = Secretary
            
        End If
        
    'Continue to ask for a name until one is given
    Loop Until Not Secretary = ""
    
    'Return the secretary
    Get_Secretary = Secretary
    
End Function

'Show Scanning User Form Subroutine:
Sub Show_Scanning_User_Form(File, Row, Last_Row, Files_To_Be_Scanned, Student_Digital_Copies)
    
    'Set the file name caption to File
    Scanning_User_Form.File_Name.Caption = File
    
    'Set the row caption to Row
    Scanning_User_Form.Row.Caption = Row
    
    'Set the last row caption to Last_Row
    Scanning_User_Form.Last_Row.Caption = Last_Row
    
    'Set the current file number caption to Row minus 1
    Scanning_User_Form.Current_File_Number.Caption = Row - 1
     
    'Set the total files caption to Last_Row minus 1
    Scanning_User_Form.Total_Files.Caption = Last_Row - 1
    
    'Set the files to be scanned caption to File_To_Be_Scanned
    Scanning_User_Form.Files_To_Be_Scanned.Caption = Files_To_Be_Scanned
    
    'Set the student digital copies caption to Student_Digital_Copies
    Scanning_User_Form.Student_Digital_Copies.Caption = Student_Digital_Copies
    
    'Open the first file in Acrobat
    Call Open_PDF(Files_To_Be_Scanned, File)
    
    'Show the Scanning_User_Form
    Scanning_User_Form.Show

End Sub

'Load Next PDF Subroutine:
Sub Load_Next_PDF(Row, Worksheet, Original_Path, Files_To_Be_Scanned_Caption)
    
    'Declare needed variables
    Dim Pop_Up_Check_Shell_Window As SHDocVw.ShellWindows
    Dim Pop_Up_Check_Object As Object
    Dim HTML_Document As HTMLDocument
    Dim HTML_Element As IHTMLElement
    Dim Computer As String
    Dim Object_WMI_Service
    Dim Items_Collection
    Dim Object_Item
    
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
            Exit For
            
        End If
        
    Next
    
    'Create a shell window
    Set Pop_Up_Check_Shell_Window = New SHDocVw.ShellWindows
    
    'Go to the next row
    Row = Row + 1
    
    'Get the file name from the new row
    Dim File
    File = Worksheet.Range("A" & Row).Value
    
    'Save the value from the last row
    Dim Last_Row
    Last_Row = Scanning_User_Form.Last_Row.Caption

    'Check to see if there are anymore files
    If File = "" Then
        
        'If there are none, close the form and exit the subroutine
        Unload Scanning_User_Form
        Exit Sub
        
    End If
    
    'Save the file name
    Scanning_User_Form.File_Name.Caption = File
    
    'Save the row
    Scanning_User_Form.Row.Caption = Row
    
    'Save the last row
    Scanning_User_Form.Last_Row.Caption = Last_Row
    
    'Save the current file number
    Scanning_User_Form.Current_File_Number.Caption = Row - 1
    
    'Save the total number of files
    Scanning_User_Form.Total_Files.Caption = Last_Row - 1
    
    'Empty the combo box
    Scanning_User_Form.Name_Files_Combo.Value = ""
    
    'Empty the student ID
    Scanning_User_Form.Student_ID_Box.Value = ""
    
    Call Open_PDF(Files_To_Be_Scanned_Caption, File)
    
    'Check if there are open windows
    If Pop_Up_Check_Shell_Window.Count Then
        
        'Loop through each open window
        For Each Pop_Up_Check_Object In Pop_Up_Check_Shell_Window
        
            'Check if the lookup pop-up is open
            If InStr(1, Pop_Up_Check_Object.LocationName, "Directory lookup/display") Then
                
                'Exit the pop up browser
                Pop_Up_Check_Object.Quit
                        
                'Delete the pop up
                Set Pop_Up_Check_Object = Nothing
            
                On Error Resume Next
                Err.Clear
                
            End If
        
        Next Pop_Up_Check_Object
        
    End If
    
End Sub

'Search ID Number Subroutine:
Sub Search_ID_Number(Student_ID, Worksheet)

    'Declare needed variables
    Dim IE_Browser
    Dim Pop_Up_Check_Shell_Window As SHDocVw.ShellWindows
    Dim Pop_Up_Check_Object As Object
    Dim HTML_Document As HTMLDocument
    Dim HTML_Element As IHTMLElement
    Dim BYU_URL As String
    Dim Second_Go
    Dim Student_Name
    Dim Pop_Up_Open As Boolean
    Dim Found As Boolean
    
    'Define Pop_Up_Open
    Pop_Up_Open = False
    
    'Define BYU URL
    BYU_URL = "https://y.byu.edu/ry/ae/prod/person/cgi/personSummary.cgi"
    
    'Create a shell window
    Set Pop_Up_Check_Shell_Window = New SHDocVw.ShellWindows
    
    'Open up internet explorer
    Set IE_Browser = CreateObject("InternetExplorer.Application")
    
    'Disable pop-ups
    IE_Browser.Silent = True
    
    'Navigate to the BYU login page
    IE_Browser.Navigate BYU_URL
    
    'Hide the browser
    IE_Browser.Visible = False
    
    'Wait for the page to load state 4
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Check if the browser was able to navigate to the correct page
    If IE_Browser.LocationURL <> BYU_URL Then
        
        'If not, output an error message, exit the workbook, quit and destroy the browser, and exit the subroutine
        MsgBox "You are not logged into your myBYU account. Please login and try again.", vbCritical, "Invalid User"
        ActiveWorkbook.Close SaveChanges:=False
        IE_Browser.Quit
        Set IE_Browser = Nothing
        Unload Scanning_User_Form
        Exit Sub
        
    End If
    
    'Save the HTML data
    Set HTML_Document = IE_Browser.document
    
    'Input the given user number into the search box
    HTML_Document.all.NAVpattern.Value = Scanning_User_Form.Student_ID_Box.Value
    
    'Search through the HTML for anchor tags
    For Each HTML_Element In HTML_Document.getElementsByTagName("a")
        
        'Skip past the first go button on the page
        If HTML_Element.innerHTML = "GO" And Second_Go = False Then
            
            Second_Go = True
            
        'Click the second go button on the page and exit the for loop
        ElseIf HTML_Element.innerHTML = "GO" And Second_Go = True Then
        
            HTML_Element.Click
            Exit For
            
        End If
        
    Next

    'Wait for the page to load
    Do
    Loop Until IE_Browser.readyState = 4
    
    'Wait for any pop ups to show
    Application.Wait (Now + TimeValue("0:00:01"))
    
    'Check if there are open windows
    If Pop_Up_Check_Shell_Window.Count Then
        
        'Loop through each open window
        For Each Pop_Up_Check_Object In Pop_Up_Check_Shell_Window
        
            Do
                'Check if the lookup pop-up is open
                If InStr(1, Pop_Up_Check_Object.LocationName, "Directory lookup/display") Then

                    'If there is a pop up, set pop up open to true
                    Pop_Up_Open = True
                    
                    'Save the pop up browser
                    Set HTML_Document = Pop_Up_Check_Object.document
                
                    'Loop through the HTML
                    For Each HTML_Element In HTML_Document.getElementsByTagName("var")
                    
                        'Check if the pop up is for an incorrect ID lookup
                        If HTML_Element.innerHTML = "J?n" Then
                        
                            'If it is set found to true
                            Found = True
                            
                        End If
                        
                    Next HTML_Element
                    
                    'There will be an error if a correct student is found from the pop up window. This line circumvents that
                    On Error GoTo Student_Found
                    
                    'Check if the incorrect ID lookup pop up was found
                    If Found = True Then
                    
                        'If so show that the pop up is open
                        Pop_Up_Open = True
                        
                        'Wait for the user to use the pop up
                        Application.Wait (Now + TimeValue("0:00:01"))
                            
                    Else
                        
                        'Show that the pop up is no long open
                        Pop_Up_Open = False
                        
                        'Exit the pop up browser
                        Pop_Up_Check_Object.Quit
                        
                        'Delete the pop up
                        Set Pop_Up_Check_Object = Nothing
                        
                        'Exit the for loop
                        Exit For
                        
                    End If
                        
                End If
            
            Loop Until Pop_Up_Open = False
            
        Next Pop_Up_Check_Object
            
    End If
    
Student_Found:

    'Wait for the page to properly load
    Application.Wait (Now + 5.787e-08)

    'Save the browser
    Set HTML_Document = IE_Browser.document
    
    'Save the student email, student name, netID, and student number, then exit the for loop
    For Each HTML_Element In HTML_Document.getElementsByTagName("input")
        
        'Check for the restore_name field
        If HTML_Element.Name = "restore_name" Then
        
            'Save that value as student name
            Student_Name = HTML_Element.Value
            On Error GoTo Error_Handling
            
        End If
        
        'Check for the restore_byu_id field
        If HTML_Element.Name = "restore_byu_id" Then
            
            'Save that value as the student ID and exit the for loop
            Student_ID = HTML_Element.Value
            On Error GoTo Error_Handling
            Student_ID = Replace(Student_ID, "-", "")
            Exit For
            
        End If

    Next
    
    'Quit and destroy the browser
    IE_Browser.Quit
    Set IE_Browser = Nothing
    
    'Search for the student folder in Student Digital Copies
    Call Search_Student_Digital_Copies(Scanning_User_Form.Student_Digital_Copies.Caption, Student_ID, Student_Name)
    Err.Clear
    
Error_Handling:

    If Err.Number <> 0 Then
        
        MsgBox "The scanning program has lost its connection to the internet. The program will now close. Please wait a moment and try again.", vbCritical, "Failed Internet Connection"
        Unload Scanning_User_Form
        ActiveWorkbook.Close SaveChanges:=False
        
    End If
    
End Sub

'Check Student Digital Copies Function:
Function Search_Student_Digital_Copies(Student_Digital_Copies, Student_ID, Student_Name)
    
    'Declare needed variables
    Dim Directory As String
    Dim Directory_Number_Only
    Dim Folder_Name As String
    Dim File_Search As String
    Dim File_Name As String
    Dim File_Number
    Dim Make_New_Folder As Integer
    Dim Row
    Dim Status
    
    'Define file number
    File_Number = 1
    
    'Get the row of the file to be scanned
    Row = Scanning_User_Form.Row.Caption
    
    'Turn of screen updating
    Application.ScreenUpdating = False
    
    'Define the directories and folder name
    Folder_Name = Student_Name & " " & Student_ID
    Directory = Student_Digital_Copies & Folder_Name
    Directory_Number_Only = Student_Digital_Copies & "*" & Student_ID
    
    'Check if the folder already exists
    If Len(Dir(Directory, vbDirectory)) > 0 Then
        
        'If it does, loop through the files in the folder
        Do
            'Set the file to search for
            File_Search = Directory & "\" & Scanning_User_Form.Name_Files_Combo.Value & File_Number & ".pdf"
            
            'Check if a file with that name already exists
            If Dir(File_Search) <> "" Then
                
                'If it does, increment the file number
                File_Number = File_Number + 1
                
            End If
            
        'Loop until no files with the same name are found
        Loop Until Dir(File_Search) = ""
        
        'Set the file name
        File_Name = Scanning_User_Form.Name_Files_Combo.Value & File_Number & ".pdf"
        
        'Copy the files into the new folder
        FileCopy Scanning_User_Form.Files_To_Be_Scanned.Caption & Cells(Row, 1).Value, File_Search
            
        'Define status
        Status = "Complete"
            
        'Record data to the spreadsheet
        ActiveWorkbook.ActiveSheet.Range("B" & Row).Value = Scanning_User_Form.Name_Files_Combo.Value & File_Number
        ActiveWorkbook.ActiveSheet.Range("C" & Row).Value = Student_ID
        ActiveWorkbook.ActiveSheet.Range("D" & Row).Value = Student_Name
        ActiveWorkbook.ActiveSheet.Range("E" & Row).Value = Status
        ActiveWorkbook.ActiveSheet.Range("E" & Row).Interior.Color = RGB(0, 255, 0)
        ActiveWorkbook.ActiveSheet.Range("F" & Row).Value = Now()
        ActiveWorkbook.ActiveSheet.Range("G" & Row).Value = Scanning_User_Form.Secretary.Caption
        ActiveWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=ActiveWorkbook.ActiveSheet.Range("H" & Row), _
            Address:=Directory, _
            ScreenTip:="Link to Student Folder", _
            TextToDisplay:=Directory
        
    'Check if the folder already exists (using the student ID only)
    ElseIf Len(Dir(Directory_Number_Only, vbDirectory)) > 0 Then
    
        'Loop through all the files in the folder
        Do
            'Set the file to search for
            File_Search = Student_Digital_Copies & Dir(Directory_Number_Only, vbDirectory) & "\" & Scanning_User_Form.Name_Files_Combo.Value & File_Number & ".pdf"
            
            'Check if a file with that name already exists
            If Dir(File_Search) <> "" Then
            
                'If it does, increment the file number
                File_Number = File_Number + 1
                
            End If
            
        'Loop until no files with the same name are found
        Loop Until Dir(File_Search) = ""
        
        'Set the file name
        File_Name = Scanning_User_Form.Name_Files_Combo.Value & File_Number & ".pdf"
        
        'Copy the files into the new folder
        FileCopy Scanning_User_Form.Files_To_Be_Scanned.Caption & Cells(Row, 1).Value, File_Search
            
        'Define status
        Status = "Complete"
            
        'Record data to the spreadsheet
        ActiveWorkbook.ActiveSheet.Range("B" & Row).Value = Scanning_User_Form.Name_Files_Combo.Value & File_Number
        ActiveWorkbook.ActiveSheet.Range("C" & Row).Value = Student_ID
        ActiveWorkbook.ActiveSheet.Range("D" & Row).Value = Student_Name
        ActiveWorkbook.ActiveSheet.Range("E" & Row).Value = Status
        ActiveWorkbook.ActiveSheet.Range("E" & Row).Interior.Color = RGB(0, 255, 0)
        ActiveWorkbook.ActiveSheet.Range("F" & Row).Value = Now()
        ActiveWorkbook.ActiveSheet.Range("G" & Row).Value = Scanning_User_Form.Secretary.Caption
        ActiveWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=ActiveWorkbook.ActiveSheet.Range("H" & Row), _
            Address:=Directory, _
            ScreenTip:="Link to Student Folder", _
            TextToDisplay:=Directory
            
    Else
    
        'Check if the user wants to make a new folder
        Make_New_Folder = MsgBox("This student does not currently have a folder in Student Digital Copies. Would you like to make one now?" & vbNewLine & vbNewLine & "Note: If you choose no, the current file will be treated as if it was skipped.", vbInformation + vbYesNo, "Create New Folder")
        
        'Check if they chose yes
        If Make_New_Folder = vbYes Then
        
            'If so make a new folder
            MkDir Directory
            
            'Set the file name
            File_Name = Directory & "\" & Scanning_User_Form.Name_Files_Combo.Value & File_Number & ".pdf"
            
            'Copy the files into the new folder
            FileCopy Scanning_User_Form.Files_To_Be_Scanned.Caption & Cells(Row, 1).Value, File_Name
            
            'Define status
            Status = "Complete"
            
            'Record data to the spreadsheet
            ActiveWorkbook.ActiveSheet.Range("B" & Row).Value = Scanning_User_Form.Name_Files_Combo.Value & File_Number
            ActiveWorkbook.ActiveSheet.Range("C" & Row).Value = Student_ID
            ActiveWorkbook.ActiveSheet.Range("D" & Row).Value = Student_Name
            ActiveWorkbook.ActiveSheet.Range("E" & Row).Value = Status
            ActiveWorkbook.ActiveSheet.Range("E" & Row).Interior.Color = RGB(0, 255, 0)
            ActiveWorkbook.ActiveSheet.Range("F" & Row).Value = Now()
            ActiveWorkbook.ActiveSheet.Range("G" & Row).Value = Scanning_User_Form.Secretary.Caption
            ActiveWorkbook.ActiveSheet.Hyperlinks.Add Anchor:=ActiveWorkbook.ActiveSheet.Range("H" & Row), _
                Address:=Directory, _
                ScreenTip:="Link to Student Folder", _
                TextToDisplay:=Directory
                
        Else
    
            'Define status
            Status = "Skipped"
    
            'Write "Skipped" in the status column and color the cell red
            ActiveWorkbook.ActiveSheet.Range("E" & Row).Value = Status
            ActiveWorkbook.ActiveSheet.Range("E" & Row).Interior.Color = RGB(255, 0, 0)
        
        End If
    
    End If
        
End Function

'Open PDF Subroutine:
Function Open_PDF(Files_To_Be_Scanned, File)
 
    'Define needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
    
    'Set the file to be opened
    PDF_Path = Files_To_Be_Scanned & File
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    
    'Create the PDF document
    Set PDF_Document = CreateObject("AcroExch.AVDoc")
    
    'Open the PDF
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'Set the focus to the PDF document
        PDF_Document.BringToFront
        
        'Maximize the document
        Call PDF_Document.Maximize(True)
        
        'Set the page view
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        
        'Set the page view of the pdf
        Call PDF_Page_View.ZoomTo(1, 50)
    
    End If
    
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Show the adobe application
    PDF_Application.Show
    
    'Set the focus to adobe acrobat pro
    AppActivate "Adobe Acrobat Pro"
    
End Function

'Delete Files Function:
Function Delete_Files(Files_To_Be_Scanned)
    
    'Declare needed variables
    Dim Response
    Dim File_System_Object As Object
    Dim Scanning_Folder
    Dim Extention
    Dim File

    'Define needed variables
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    Set Scanning_Folder = File_System_Object.GetFolder(Files_To_Be_Scanned)
    Extention = ""
    
    'Ask if the user wants to empty the To_Be_Scanned Folder
    Response = MsgBox("Would you like to empty the To_Be_Scanned folder now?", vbInformation + vbYesNo, "Empty Folder")
    
    'Check if the user said yes
    If Response = vbYes Then
    
        'Loop through each file
        For Each File In Scanning_Folder.Files
    
            'Save the file extension
            Extention = Right$(File.Path, Len(File.Path) - InStrRev(File.Path, "."))
        
            'Check for pdf files
            If Extention = "pdf" Then
        
                'If found, delete the file
                Kill File
            
            End If
        
        Next
        
        'Delete File_System_Object and Scanning_Folder
        Set File_System_Object = Nothing
        Set Scanning_Folder = Nothing
        
     End If
        
End Function

'Close Program Function:
Function Close_Program()

    'Declare needed variables
    Dim Response

    'Ask if the user wants to close the program
    Response = MsgBox("Would you like to close this excel workbook now?", vbInformation + vbYesNo, "Close Workbook")
    
    'Check if the user said yes
    If Response = vbYes Then
    
        'If so close the workbook
        ActiveWorkbook.Close SaveChanges:=False
        
    End If

End Function
