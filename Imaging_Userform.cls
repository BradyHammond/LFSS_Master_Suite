'=================== IMAGING FORM ==================='
'Created By: Brady Hammond
'Created: 07/06/16
'===================================================='

'Initialize User Form Subroutine:
Private Sub UserForm_Initialize()

    'Reset all fields
    Imaging_User_Form.Student_Information_Input.Value = ""
    Imaging_User_Form.Student_File_Label.Caption = "Student File: "
    Imaging_User_Form.Page_Selection.Pages(2).Visible = True
    Imaging_User_Form.Page_Selection.Pages(0).Visible = True
    Imaging_User_Form.Page_Selection.Pages(1).Visible = False
    Imaging_User_Form.Height = 110
    Imaging_User_Form.Page_Selection.Height = 84
    Imaging_User_Form.Student_File_Label.Caption = "Student File: "
    Imaging_User_Form.Submit_Button_II.Top = 48
    Imaging_User_Form.Cancel_Button_III.Top = 48
    Imaging_User_Form.File_Label_1.Caption = "File"
    Imaging_User_Form.File_Label_1.Visible = False
    Imaging_User_Form.File_Label_2.Caption = "File"
    Imaging_User_Form.File_Label_2.Visible = False
    Imaging_User_Form.File_Label_3.Caption = "File"
    Imaging_User_Form.File_Label_3.Visible = False
    Imaging_User_Form.File_Label_4.Caption = "File"
    Imaging_User_Form.File_Label_4.Visible = False
    Imaging_User_Form.File_Label_5.Caption = "File"
    Imaging_User_Form.File_Label_5.Visible = False
    Imaging_User_Form.File_Label_6.Caption = "File"
    Imaging_User_Form.File_Label_6.Visible = False
    Imaging_User_Form.File_Label_7.Caption = "File"
    Imaging_User_Form.File_Label_7.Visible = False
    Imaging_User_Form.File_Label_8.Caption = "File"
    Imaging_User_Form.File_Label_8.Visible = False
    Imaging_User_Form.File_Label_9.Caption = "File"
    Imaging_User_Form.File_Label_9.Visible = False
    Imaging_User_Form.File_Label_10.Caption = "File"
    Imaging_User_Form.File_Label_10.Visible = False
    Imaging_User_Form.File_Label_11.Caption = "File"
    Imaging_User_Form.File_Label_11.Visible = False
    Imaging_User_Form.File_Label_12.Caption = "File"
    Imaging_User_Form.File_Label_12.Visible = False
    Imaging_User_Form.File_Label_13.Caption = "File"
    Imaging_User_Form.File_Label_13.Visible = False
    Imaging_User_Form.File_Label_14.Caption = "File"
    Imaging_User_Form.File_Label_14.Visible = False
    Imaging_User_Form.File_Label_15.Caption = "File"
    Imaging_User_Form.File_Label_15.Visible = False
    Imaging_User_Form.File_Label_16.Caption = "File"
    Imaging_User_Form.File_Label_16.Visible = False
    Imaging_User_Form.File_Label_17.Caption = "File"
    Imaging_User_Form.File_Label_17.Visible = False
    Imaging_User_Form.File_Label_18.Caption = "File"
    Imaging_User_Form.File_Label_18.Visible = False
    Imaging_User_Form.File_Label_19.Caption = "File"
    Imaging_User_Form.File_Label_19.Visible = False
    Imaging_User_Form.File_Label_20.Caption = "File"
    Imaging_User_Form.File_Label_20.Visible = False
    Imaging_User_Form.File_Label_21.Caption = "File"
    Imaging_User_Form.File_Label_21.Visible = False
    Imaging_User_Form.File_Label_22.Caption = "File"
    Imaging_User_Form.File_Label_22.Visible = False
    Imaging_User_Form.File_Label_23.Caption = "File"
    Imaging_User_Form.File_Label_23.Visible = False
    Imaging_User_Form.File_Label_24.Caption = "File"
    Imaging_User_Form.File_Label_24.Visible = False
    Imaging_User_Form.File_Label_25.Caption = "File"
    Imaging_User_Form.File_Label_25.Visible = False
    Imaging_User_Form.File_Label_26.Caption = "File"
    Imaging_User_Form.File_Label_26.Visible = False
    
End Sub

'Submit Button Enter Subroutine:
Private Sub Submit_Button_I_Enter()
    
    'Check if there was a value entered
    If Imaging_User_Form.Student_Information_Input.Value = "" Then
        
        'If not, output an error message
        MsgBox "You have not entered a student ID number. Please enter one and try again.", vbInformation, "Invalid Input"
    
    Else
        
        'Check if the user input a number
        If IsNumeric(Imaging_User_Form.Student_Information_Input.Value) = True Then
        
            'If so, check if the number is 9 characters long
            If Len(Imaging_User_Form.Student_Information_Input.Value) <> 9 Then
            
                'If not, output an error message
                MsgBox "The number you entered is not the correct length. Please try again.", vbInformation, "Invalid Input"
            
            Else
                
                'Search for the student file
                Search_Student_Digital_Copies_Imaging (Imaging_User_Form.Student_Information_Input.Value)
                
            End If
            
        Else
        
            'Output an error message
            MsgBox "The input you have provided is not a number. Please correct this issue and try again.", vbInformation, "Invalid Input"
            
        End If
        
    End If
    
End Sub

'Submit Button Click Subroutine:
Private Sub Submit_Button_II_Click()

    'Declare needed variables
    Dim Control As Control
    Dim Directory_Name As String
    Dim File_System_Object As Object
    Dim Directory As Folder
    Dim Folder As FileDialog
    Dim Selected_Item As String
    Dim Destination_Path As String
    
    'Define which directory to use and create a file system object
    Directory_Name = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "")
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    Set Directory = File_System_Object.GetFolder(Directory_Name)
    
    'Check if the directory is empty
    If Directory.Files.Count = 0 Then
    
        'If so, delete the direcotry
        File_System_Object.DeleteFolder Directory_Name, True
    
    Else
    
        'Output an instruction message
        MsgBox "Please select a location to place this student's folder.", vbInformation, "Folder Select"
        
        'Open a file diolog picker
        Set Folder = Application.FileDialog(msoFileDialogFolderPicker)
        
        'Set the attributes for the folder
        With Folder
        
            .Title = "Folder Select"
            .AllowMultiSelect = Flase
            .InitialFileName = "M:\To Be Imaged"
            .Show
            Selected_Item = .SelectedItems(1)
            
        End With
        
        'Check if the directory name ends in a backslash
        If Right(Directory_Name, 1) = "\" Then
        
            'If so, remove it from the directory name
            Directory_Name = Left(Directory_Name, Len(Directory_Name) - 1)
        
        End If
        
        'Check if the selected folder has a backslash in it
        If Right(Selected_Item, 1) = "\" Then
        
            'If so, set the destination path accordingly
            Destination_Path = Selected_Item & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "")
        
        Else
        
            'Set the destination path
            Destination_Path = Selected_Item & "\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "")
            
        End If
        
        'Check if the student's folder already exists in the destination folder
        If File_System_Object.FolderExists(Destination_Path) = True Then
        
            'If so, output an error message and exit the subroutine
            MsgBox "The folder: " & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & " already exists in this folder.", vbInformation, "File Already Exists"
            Exit Sub
            
        End If
        
        'Check if the student's folder exists
        If File_System_Object.FolderExists(Directory_Name) = False Then
        
            'If not, output an error message and exit the subroutine
            MsgBox Directory_Name & " does not exist.", vbCritical, "Invalid Directory"
            Exit Sub
            
        End If
        
        'Move the student folder
        File_System_Object.MoveFolder Source:=Directory_Name, Destination:=Destination_Path
        
        'Delete the dialog picker
        Set Folder = Nothing
        
    End If
        
    'Reset all of the userform properties
    Imaging_User_Form.Page_Selection.Pages(2).Visible = True
    Imaging_User_Form.Page_Selection.Pages(0).Visible = True
    Imaging_User_Form.Page_Selection.Pages(1).Visible = False
    Imaging_User_Form.Height = 110
    Imaging_User_Form.Page_Selection.Height = 84
    Imaging_User_Form.Student_File_Label.Caption = "Student File: "
    Imaging_User_Form.Submit_Button_II.Top = 48
    Imaging_User_Form.Cancel_Button_III.Top = 48
    Imaging_User_Form.File_Label_1.Caption = "File"
    Imaging_User_Form.File_Label_1.Visible = False
    Imaging_User_Form.File_Label_2.Caption = "File"
    Imaging_User_Form.File_Label_2.Visible = False
    Imaging_User_Form.File_Label_3.Caption = "File"
    Imaging_User_Form.File_Label_3.Visible = False
    Imaging_User_Form.File_Label_4.Caption = "File"
    Imaging_User_Form.File_Label_4.Visible = False
    Imaging_User_Form.File_Label_5.Caption = "File"
    Imaging_User_Form.File_Label_5.Visible = False
    Imaging_User_Form.File_Label_6.Caption = "File"
    Imaging_User_Form.File_Label_6.Visible = False
    Imaging_User_Form.File_Label_7.Caption = "File"
    Imaging_User_Form.File_Label_7.Visible = False
    Imaging_User_Form.File_Label_8.Caption = "File"
    Imaging_User_Form.File_Label_8.Visible = False
    Imaging_User_Form.File_Label_9.Caption = "File"
    Imaging_User_Form.File_Label_9.Visible = False
    Imaging_User_Form.File_Label_10.Caption = "File"
    Imaging_User_Form.File_Label_10.Visible = False
    Imaging_User_Form.File_Label_11.Caption = "File"
    Imaging_User_Form.File_Label_11.Visible = False
    Imaging_User_Form.File_Label_12.Caption = "File"
    Imaging_User_Form.File_Label_12.Visible = False
    Imaging_User_Form.File_Label_13.Caption = "File"
    Imaging_User_Form.File_Label_13.Visible = False
    Imaging_User_Form.File_Label_14.Caption = "File"
    Imaging_User_Form.File_Label_14.Visible = False
    Imaging_User_Form.File_Label_15.Caption = "File"
    Imaging_User_Form.File_Label_15.Visible = False
    Imaging_User_Form.File_Label_16.Caption = "File"
    Imaging_User_Form.File_Label_16.Visible = False
    Imaging_User_Form.File_Label_17.Caption = "File"
    Imaging_User_Form.File_Label_17.Visible = False
    Imaging_User_Form.File_Label_18.Caption = "File"
    Imaging_User_Form.File_Label_18.Visible = False
    Imaging_User_Form.File_Label_19.Caption = "File"
    Imaging_User_Form.File_Label_19.Visible = False
    Imaging_User_Form.File_Label_20.Caption = "File"
    Imaging_User_Form.File_Label_20.Visible = False
    Imaging_User_Form.File_Label_21.Caption = "File"
    Imaging_User_Form.File_Label_21.Visible = False
    Imaging_User_Form.File_Label_22.Caption = "File"
    Imaging_User_Form.File_Label_22.Visible = False
    Imaging_User_Form.File_Label_23.Caption = "File"
    Imaging_User_Form.File_Label_23.Visible = False
    Imaging_User_Form.File_Label_24.Caption = "File"
    Imaging_User_Form.File_Label_24.Visible = False
    Imaging_User_Form.File_Label_25.Caption = "File"
    Imaging_User_Form.File_Label_25.Visible = False
    Imaging_User_Form.File_Label_26.Caption = "File"
    Imaging_User_Form.File_Label_26.Visible = False
    
    Imaging_User_Form.Page_Selection.Value = 0
    
End Sub

'Cancel Button Click Subroutine:
Private Sub Cancel_Button_I_Click()

    'Unload the user form
    Unload Imaging_User_Form
    
End Sub

'Cancel Button Click Subroutine:
Private Sub Cancel_Button_II_Click()

    'Unload the user form
    Unload Imaging_User_Form
    
End Sub

'Cancel Button Click Subroutine:
Private Sub Cancel_Button_III_Click()

    'Reset all fields
    Imaging_User_Form.Student_Information_Input.Value = ""
    Imaging_User_Form.Student_File_Label.Caption = "Student File: "
    Imaging_User_Form.Page_Selection.Pages(0).Visible = True
    Imaging_User_Form.Page_Selection.Pages(1).Visible = False
    Imaging_User_Form.Height = 110
    Imaging_User_Form.Page_Selection.Height = 84
    Imaging_User_Form.Student_File_Label.Caption = "Student File: "
    Imaging_User_Form.Submit_Button_II.Top = 48
    Imaging_User_Form.Cancel_Button_III.Top = 48
    Imaging_User_Form.File_Label_1.Caption = "File"
    Imaging_User_Form.File_Label_1.Visible = False
    Imaging_User_Form.File_Label_2.Caption = "File"
    Imaging_User_Form.File_Label_2.Visible = False
    Imaging_User_Form.File_Label_3.Caption = "File"
    Imaging_User_Form.File_Label_3.Visible = False
    Imaging_User_Form.File_Label_4.Caption = "File"
    Imaging_User_Form.File_Label_4.Visible = False
    Imaging_User_Form.File_Label_5.Caption = "File"
    Imaging_User_Form.File_Label_5.Visible = False
    Imaging_User_Form.File_Label_6.Caption = "File"
    Imaging_User_Form.File_Label_6.Visible = False
    Imaging_User_Form.File_Label_7.Caption = "File"
    Imaging_User_Form.File_Label_7.Visible = False
    Imaging_User_Form.File_Label_8.Caption = "File"
    Imaging_User_Form.File_Label_8.Visible = False
    Imaging_User_Form.File_Label_9.Caption = "File"
    Imaging_User_Form.File_Label_9.Visible = False
    Imaging_User_Form.File_Label_10.Caption = "File"
    Imaging_User_Form.File_Label_10.Visible = False
    Imaging_User_Form.File_Label_11.Caption = "File"
    Imaging_User_Form.File_Label_11.Visible = False
    Imaging_User_Form.File_Label_12.Caption = "File"
    Imaging_User_Form.File_Label_12.Visible = False
    Imaging_User_Form.File_Label_13.Caption = "File"
    Imaging_User_Form.File_Label_13.Visible = False
    Imaging_User_Form.File_Label_14.Caption = "File"
    Imaging_User_Form.File_Label_14.Visible = False
    Imaging_User_Form.File_Label_15.Caption = "File"
    Imaging_User_Form.File_Label_15.Visible = False
    Imaging_User_Form.File_Label_16.Caption = "File"
    Imaging_User_Form.File_Label_16.Visible = False
    Imaging_User_Form.File_Label_17.Caption = "File"
    Imaging_User_Form.File_Label_17.Visible = False
    Imaging_User_Form.File_Label_18.Caption = "File"
    Imaging_User_Form.File_Label_18.Visible = False
    Imaging_User_Form.File_Label_19.Caption = "File"
    Imaging_User_Form.File_Label_19.Visible = False
    Imaging_User_Form.File_Label_20.Caption = "File"
    Imaging_User_Form.File_Label_20.Visible = False
    Imaging_User_Form.File_Label_21.Caption = "File"
    Imaging_User_Form.File_Label_21.Visible = False
    Imaging_User_Form.File_Label_22.Caption = "File"
    Imaging_User_Form.File_Label_22.Visible = False
    Imaging_User_Form.File_Label_23.Caption = "File"
    Imaging_User_Form.File_Label_23.Visible = False
    Imaging_User_Form.File_Label_24.Caption = "File"
    Imaging_User_Form.File_Label_24.Visible = False
    Imaging_User_Form.File_Label_25.Caption = "File"
    Imaging_User_Form.File_Label_25.Visible = False
    Imaging_User_Form.File_Label_26.Caption = "File"
    Imaging_User_Form.File_Label_26.Visible = False
    
    'Unload the user form
    Unload Imaging_User_Form
    
End Sub

'View Button Click Subroutine:
Private Sub View_Button_1_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_1
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_2_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_2
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_3_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_3
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_4_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_4
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_5_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_5
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_6_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_6
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_7_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_7
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub


'View Button Click Subroutine:
Private Sub View_Button_8_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_8
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_9_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_9
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_10_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_10
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_11_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_11
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_12_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_12
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_13_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_13
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_14_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_14
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_15_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_15
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_16_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_16
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_17_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_17
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_18_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_18
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_19_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_19
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_20_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_20
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_21_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_21
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_22_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_22
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_23_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_23
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_24_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_24
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_25_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_25
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'View Button Click Subroutine:
Private Sub View_Button_26_Click()

    'Declare needed variables
    Dim PDF_Application As AcroApp
    Dim PDF_Document As AcroAVDoc
    Dim PDF_Page_View As AcroAVPageView
    Dim PDF_Path As String
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
    
    'Define the PDF path
    PDF_Path = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_26
    
    'Open Acrobat
    Set PDF_Application = CreateObject("AcroExch.App")
    Set PDF_Document = CreateObject("AcroExch.AVDoc")

    'Check if the PDF opens
    If PDF_Document.Open(PDF_Path, "") = True Then
        
        'If so, set up the correct view
        PDF_Document.BringToFront
        Call PDF_Document.Maximize(True)
        Set PDF_Page_View = PDF_Document.GetAVPageView()
        Call PDF_Page_View.ZoomTo(1, 50)

    Else

        'Output an error message
        MsgBox "Acrobat could not locate the correct file", vbCritical, "File Error"

    End If

    'Delete used variables
    Set PDF_Document = Nothing
    Set PDF_Application = Nothing
    
    On Error Resume Next
    Err.Clear
    
    'Activate Acrobat
    AppActivate "Adobe Acrobat Pro"

End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_1_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_1.Visible = False
    Imaging_User_Form.View_Button_1.Visible = False
    Imaging_User_Form.Delete_Button_1.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_1
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_2_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_2.Visible = False
    Imaging_User_Form.View_Button_2.Visible = False
    Imaging_User_Form.Delete_Button_2.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_2
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_3_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_3.Visible = False
    Imaging_User_Form.View_Button_3.Visible = False
    Imaging_User_Form.Delete_Button_3.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_3
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_4_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_4.Visible = False
    Imaging_User_Form.View_Button_4.Visible = False
    Imaging_User_Form.Delete_Button_4.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_4
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_5_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_5.Visible = False
    Imaging_User_Form.View_Button_5.Visible = False
    Imaging_User_Form.Delete_Button_5.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_5
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_6_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_6.Visible = False
    Imaging_User_Form.View_Button_6.Visible = False
    Imaging_User_Form.Delete_Button_6.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_6
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_7_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_7.Visible = False
    Imaging_User_Form.View_Button_7.Visible = False
    Imaging_User_Form.Delete_Button_7.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_7
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_8_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_8.Visible = False
    Imaging_User_Form.View_Button_8.Visible = False
    Imaging_User_Form.Delete_Button_8.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_8
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_9_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_9.Visible = False
    Imaging_User_Form.View_Button_9.Visible = False
    Imaging_User_Form.Delete_Button_9.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_9
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_10_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_10.Visible = False
    Imaging_User_Form.View_Button_10.Visible = False
    Imaging_User_Form.Delete_Button_10.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_10
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_11_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_11.Visible = False
    Imaging_User_Form.View_Button_11.Visible = False
    Imaging_User_Form.Delete_Button_11.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_11
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_12_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_12.Visible = False
    Imaging_User_Form.View_Button_12.Visible = False
    Imaging_User_Form.Delete_Button_12.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_12
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_13_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_13.Visible = False
    Imaging_User_Form.View_Button_13.Visible = False
    Imaging_User_Form.Delete_Button_13.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_13
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_14_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_14.Visible = False
    Imaging_User_Form.View_Button_14.Visible = False
    Imaging_User_Form.Delete_Button_14.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_14
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_15_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_15.Visible = False
    Imaging_User_Form.View_Button_15.Visible = False
    Imaging_User_Form.Delete_Button_15.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_15
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_16_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_16.Visible = False
    Imaging_User_Form.View_Button_16.Visible = False
    Imaging_User_Form.Delete_Button_16.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_16
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_17_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_17.Visible = False
    Imaging_User_Form.View_Button_17.Visible = False
    Imaging_User_Form.Delete_Button_17.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_17
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_18_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_18.Visible = False
    Imaging_User_Form.View_Button_18.Visible = False
    Imaging_User_Form.Delete_Button_18.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_18
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_19_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_19.Visible = False
    Imaging_User_Form.View_Button_19.Visible = False
    Imaging_User_Form.Delete_Button_19.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_19
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_20_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_20.Visible = False
    Imaging_User_Form.View_Button_20.Visible = False
    Imaging_User_Form.Delete_Button_20.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_20
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_21_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_21.Visible = False
    Imaging_User_Form.View_Button_21.Visible = False
    Imaging_User_Form.Delete_Button_21.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_21
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_22_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_22.Visible = False
    Imaging_User_Form.View_Button_22.Visible = False
    Imaging_User_Form.Delete_Button_22.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_22
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_23_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_23.Visible = False
    Imaging_User_Form.View_Button_23.Visible = False
    Imaging_User_Form.Delete_Button_23.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_23
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_24_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_24.Visible = False
    Imaging_User_Form.View_Button_24.Visible = False
    Imaging_User_Form.Delete_Button_24.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_24
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_25_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_25.Visible = False
    Imaging_User_Form.View_Button_25.Visible = False
    Imaging_User_Form.Delete_Button_25.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_25
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Delete Button Click Subroutine:
Private Sub Delete_Button_26_Click()

    'Declare needed variables
    Dim Delete_File
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
    
    'Rehide the label and buttons
    Imaging_User_Form.File_Label_26.Visible = False
    Imaging_User_Form.View_Button_26.Visible = False
    Imaging_User_Form.Delete_Button_26.Visible = False
    
    'Set the file to be deleted
    Delete_File = "M:\Student Digital Copies\" & Replace(Imaging_User_Form.Student_File_Label, "Student File: ", "") & "\" & File_Label_26
    
    'Check if the file exists
    If Dir(Delete_File) <> "" Then
    
        'If so, make sure the file is deletable
        SetAttr Delete_File, vbNormal
    
        'Delete the file
        Kill Delete_File

    Else
        
        'Output an error message
        MsgBox "The file you are trying to delete could not be found.", vbInformation, "Missing File"
    
    End If
    
End Sub

'Merge Button Subroutine:
Private Sub Merge_Button_Click()
    Call Merge_PDFs
End Sub
