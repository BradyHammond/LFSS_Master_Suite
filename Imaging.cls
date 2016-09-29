'========================= IMAGING ========================='
'Created By: Brady Hammond
'Created: 05/23/16
'==========================================================='

'Imaging Subroutine:
Sub IMAGING_PROGRAM()

    'Show the user form
    Imaging_User_Form.Show
    
End Sub

'Merge PDF's Subroutine:
Sub Merge_PDFs()

    'Declare needed variables
    Dim Select_File As FileDialog
    Dim File_System_Object As Object
    Dim Directory As Folder
    Dim Selected_Item As String
    Dim Merge_Folder As String
    Dim Folder As Folder
    Dim File As File
    Dim Path_Name As String
    Dim File_List As String
    Dim File_Array As Variant
    Dim Acrobat As New Acrobat.AcroApp
    Dim Part_Documents() As Acrobat.AcroPDDoc
    Dim Number_of_Pages As Long
    Dim Page_Number As Long
    Dim i As Long
    Dim File_Counter As Long
    
    'Output an instruction message
    MsgBox "Please select the folder you would like to merge documents from.", vbInformation, "Folder Selection"
    
    'Create the file selection window
    Set Select_File = Application.FileDialog(msoFileDialogFolderPicker)
    
    'Set the attributes for the selection window
    With Select_File
    
        .Title = "Folder Select"
        .AllowMultiSelect = False
        .InitialFileName = "M:\To Be Imaged"
        .Show
        Selected_Item = .SelectedItems(1)
        
    End With
    
    'Check if there is a trailing backslash in the selected item path
    If Right(Selected_Item, 1) = "\" Then
        
        'If so, define the merge folder path accordingly
        Merge_Folder = Selected_Item & "Merge Folder"
        
    Else
    
        'Otherwise, define the merge folder path accordingly
        Merge_Folder = Selected_Item & "\Merge Folder"
    
    End If
    
    'Check if the merge folder already exists
    If Len(Dir(Merge_Folder, vbDirectory)) = 0 Then
    
        'If not, create the merge folder
        MkDir Merge_Folder
        
    Else
        
        'Otherwise output an error message and exit the subroutine
        MsgBox "There seems to already be a merged pdf document from this folder. If you feel this is incorrect or would like to remerge the documents from this folder, please delete the folder titled " & Chr(34) & "Merge Folder" & Chr(34) & " and try again.", vbInformation, "Folder Previously Merged"
        Exit Sub
        
    End If
    
    'Create the file system object and set the directory
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    Set Directory = File_System_Object.GetFolder(Selected_Item)
    
    'Define the file counter
    File_Counter = 1
    
    'Loop through each folder in the directory
    For Each Folder In Directory.SubFolders
    
        'Check if the folder is the merge folder
        If Folder.Name <> "Merge Folder" Then
        
            'If not, loop through each file in the subfolder
            For Each File In Folder.Files
                
                'Check if the file is a pdf
                If File.Name Like "*.pdf" Then
                
                    'If so, edit its attributes, copy the file into the merge folder, and increment the file counter
                    SetAttr File, vbNormal
                    File_System_Object.CopyFile File.Path, Merge_Folder & "\" & "File" & File_Counter & ".pdf"
                    File_Counter = File_Counter + 1
                    
                End If
                
            Next File
            
        End If
        
    Next Folder
    
    'Redefine the directory as the merge folder
    Set Directory = File_System_Object.GetFolder(Merge_Folder)
    
    'Check if the merge folder has a trailing backslash
    If Right(Merge_Folder, 1) <> "\" Then
    
        'If not, add the backslash
        Merge_Folder = Merge_Folder & "\"
        
    End If

    'Loop through each file in the directory
    For Each File In Directory.Files
    
        'Add each file to a list
        File_List = File.ShortName & "," & File_List

    Next File
    
    'Remove the trailing comma at the end of the list
    File_List = Left(File_List, Len(File_List) - 1)
    
    'Add each file from the list to an array
    File_Array = Split(File_List, ",")
    
    'Redeclare part documents to have the correct array dimensions
    ReDim Part_Documents(0 To UBound(File_Array))
    
    On Error GoTo Exit_Subroutine
    
    'Checl if a merged file already exists
    If Len(Dir(Merge_Folder & "Merged_File.pdf")) Then
        
        'If so, delete it
        Kill Merge_Folder & "Merged_File.pdf"
    
    End If
    
    'Loop through the array
    For i = 0 To UBound(File_Array)
    
        'Check if each file exists
        If Dir(Merge_Folder & Trim(File_Array(i))) = "" Then
        
            'If not, output an error message and exit the subroutine
            MsgBox "File not found" & vbLf & Merge_Folder & File_Array(i), vbCritical, "Canceled"
            Exit Sub
            
        End If
        
        'Open part documents as an acrobat document
        Set Part_Documents(i) = CreateObject("AcroExch.PDDoc")
        Part_Documents(i).Open Merge_Folder & Trim(File_Array(i))
        
        'Check the current file
        If i Then
        
            'Get the number of pages
            Number_of_Pages = Part_Documents(i).GetNumPages()
            
            'Check if the page was inserted correctly
            If Not Part_Documents(0).InsertPages(Page_Number - 1, Part_Documents(i), 0, Number_of_Pages, True) Then
                
                'If not, output an error message and exit the subroutine
                MsgBox "Cannot insert pages of" & vbLf & Merge_Folder & File_Array(i), vbCritical, "Canceled"
                Exit Sub
                
            End If
            
            'Get the page number
            Page_Number = Page_Number + Number_of_Pages
            
            'Close the document
            Part_Documents(i).Close
            Set Part_Documents(i) = Nothing
        
        Else
            
            'Get the page number
            Page_Number = Part_Documents(0).GetNumPages()
            
        End If
        
    Next
    
    'Check if i exceeds the array size
    If i > UBound(File_Array) Then
    
        'If so, check if the file was saved
        If Not Part_Documents(0).Save(PDSaveFull, Merge_Folder & "Merged_File.pdf") Then
        
            'If not, output an error message and exit the subroutine
            MsgBox "Cannot save the resulting document" & vbLf & Mege_Folder & "Merged_File.pdf", vbCritical, "Canceled"
            Exit Sub
            
        End If
        
    End If
                 
Exit_Subroutine:

    'Check if there was an error
    If Err Then
    
        'If so output an error message and exit the subroutine
        MsgBox Err.Description, vbCritical, "Error: " & Err.Number
        Exit Sub
        
    'Check if i exceeds the array size
    ElseIf i > UBound(File_Array) Then
    
        'Redefine the directory
        Set Directory = File_System_Object.GetFolder(Merge_Folder)
        
        'Loop through each file in the directory
        For Each File In Directory.Files
        
            'Check if the file was a copied file
            If File.Name Like "File*.pdf" Then
            
                'If so, edit the file attributes and delete the file
                SetAttr File, vbNormal
                Kill File
                
            End If
        
        Next File
        
        'Output a success message
        MsgBox vbLf & Merge_Folder & "Merged_File.pdf" & " was succesfully created.", vbInformation, "Merge Success"
        
    End If
    
    'Check if the part documents still exists
    If Not Part_Documents(0) Is Nothing Then
        
        'If so, close it
        Part_Documents(0).Close
        
    End If
    
    'Set part documents to nothing
    Set Part_Documents(0) = Nothing
    
    'Exit acrobat and set it to nothing
    Acrobat.Exit
    Set Acrobat = Nothing
        
End Sub

'Search Student Digital Copies Subroutine:
Sub Search_Student_Digital_Copies_Imaging(Student_ID)

    'Declare needed variables
    Dim File_System_Object As Object
    Dim Directory As Folder
    Dim Folder
    Dim File As File
    Dim Found As Boolean
    Dim Top_Counter As Integer
    Dim Control_Counter As Integer
    Dim Line_Number As Long
    Dim Sub_Name As String
    Dim Program As String
    Dim Control As Control
    
    'Clear the input field on the userform
    Imaging_User_Form.Student_Information_Input.Value = ""
    
    'Create the file system object
    Set File_System_Object = CreateObject("Scripting.FileSystemObject")
    
    'Set the default for the found variable
    Found = False
    
    'Set the default for the top counter and label counter
    Top_Counter = 40
    Control_Counter = 1
    
    'Search for the student folder
    Folder = Dir("M:\Student Digital Copies\*" & Student_ID, vbDirectory)
    
    'Check if the folder was found
    If Len(Folder) <> 0 Then
        
        'Set the directory to the student folder
        Set Directory = File_System_Object.GetFolder("M:\Student Digital Copies\" & Folder)
                
        'If so, show the file name
        Imaging_User_Form.Student_File_Label.Caption = Imaging_User_Form.Student_File_Label.Caption & Directory.Name
            
        'Loop through each file in the directory
        For Each File In Directory.Files
            
            'Loop through each control on the form
            For Each Control In Imaging_User_Form.Controls
    
                'Check for the correct label
                If Control.Name = "File_Label_" & Control_Counter Then
                    
                    'If found, set the caption and make it visible
                    Control.Caption = File.Name
                    Control.Visible = True
                    
                'Check for the correct view button
                ElseIf Control.Name = "View_Button_" & Control_Counter Then
                    
                    'If found, make it visible
                    Control.Visible = True
                        
                'Check for the correct delete button
                ElseIf Control.Name = "Delete_Button_" & Control_Counter Then
                    
                    'If found, make it visible
                    Control.Visible = True
                         
                End If
                    
            Next
                
            'Edit the userform properties
            Imaging_User_Form.Height = Imaging_User_Form.Height + 20
            Imaging_User_Form.Page_Selection.Height = Imaging_User_Form.Page_Selection.Height + 20
            Imaging_User_Form.Submit_Button_II.Top = Imaging_User_Form.Submit_Button_II.Top + 20
            Imaging_User_Form.Cancel_Button_III.Top = Imaging_User_Form.Cancel_Button_III.Top + 20
                
            'Increment the necessary counters
            Top_Counter = Top_Counter + 20
            Control_Counter = Control_Counter + 1
                
        Next
            
        'Show the page
        Imaging_User_Form.Page_Selection.Pages(0).Visible = False
        Imaging_User_Form.Page_Selection.Pages(2).Visible = False
        Imaging_User_Form.Page_Selection.Pages(1).Visible = True
        Imaging_User_Form.Page_Selection.Value = 1
            
        'Set found to true and exit the for loop
        Found = True
    
    Else
    
        Found = False
        
    End If
    
    'Check if there are too many files in the folder
    If Control_Counter > 26 Then
        
        'If so, output an error message and exit the for loop
        MsgBox "This folder has more files than the display limit.", vbInformation, "File Limit Exceeded"

            
    End If
    
    'Check if the student folder was found
    If Found <> True Then
            
        'If not output an error message
        MsgBox "The student you were looking for was not found.", vbInformation, "Student Not Found"
            
    End If
    
End Sub
