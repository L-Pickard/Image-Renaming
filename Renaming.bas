Option Explicit

'===================================================================================================================
'Project: Master Image Renaming
'Language: VBA
'Author: Leo Pickard
'Date: March 2025
'Version: 2.0

'This vba file contains subroutines and functions which are used to process image files into the required file name
'formats. The main subroutine which is executed is process_images, the other bolocks are called by this subroutine.
'===================================================================================================================

'Below is a global variable used to collate the execution log entries

Dim collect_log     As Collection

'Below is the subroutine which writes the log event collection data to the execution log worksheet table.

Sub write_logs_to_table()
    
    Dim sh_log      As Worksheet
    Dim tbl_log     As ListObject
    Dim log_array() As Variant
    Dim i           As Long, j As Long
    Dim num_rows    As Long
    
    Set sh_log = log
    Set tbl_log = sh_log.ListObjects(1)
    
    If collect_log Is Nothing Or collect_log.Count = 0 Then
        Exit Sub
    End If
    
    If tbl_log.ListRows.Count > 0 Then tbl_log.DataBodyRange.Delete
    
    num_rows = collect_log.Count
    ReDim log_array(1 To num_rows, 1 To 8)
    
    For i = 1 To num_rows
        Dim log_entry As Variant
        log_entry = collect_log(i)
        For j = 1 To 8
            log_array(i, j) = log_entry(j - 1)
        Next j
    Next i
    
    If num_rows > 0 Then
        tbl_log.ListRows.Add AlwaysInsert:=True
        tbl_log.DataBodyRange.Resize(num_rows, 8).Value = log_array
    End If
    
    Set collect_log = New Collection
    
End Sub

'The below subroutine is used to insert a new log record into the global collection.

Sub insert_log_row(action As String, status As String, Optional sub_dir As String = "", _
    Optional input_file As String = "", Optional output_file As String = "", _
    Optional error_message As String = "")
    
    If collect_log Is Nothing Then Set collect_log = New Collection
    
    Dim log_record  As Variant
    log_record = Array(Int(Now), Format(Now, "HH:MM:SS"), input_file, sub_dir, output_file, action, status, error_message)
    
    collect_log.Add log_record
    
End Sub

'The below function handles the file copy and rename process. It is repeatedly called by the process-image subroutine.

Function copy_rename_files(fso As Object, size_name As String) As Boolean
    
    Dim sh_controls As Worksheet
    Dim sh_rename   As Worksheet
    Dim tbl_rename  As ListObject
    Dim data_array  As Variant
    Dim i           As Long
    Dim folder_name As String
    Dim sub_folder  As String
    Dim image_folder As String
    Dim output_folder As String
    Dim input_file  As String, output_file As String
    Dim input_path  As String, output_path As String
    Dim error_message As String
    Dim answer      As Integer
    
    Set sh_rename = rename
    Set sh_controls = controls
    Set tbl_rename = sh_rename.ListObjects(1)
    
    output_folder = sh_controls.Range("Output_Folder").Value
    image_folder = sh_controls.Range("Image_Folder").Value
    
    folder_name = Replace(size_name, "/", "~")
    
    sh_rename.Range("Rename_Selection").Value = size_name
    
    Call insert_log_row("Set renaming data worksheet To rename " & size_name & " images.", "Success", size_name)
    
    sh_rename.Calculate
    
    Call insert_log_row("Calculate renaming images worksheet ready For copy files operations.", "Success", size_name)
    
    data_array = tbl_rename.ListColumns("Name").DataBodyRange.Resize(, 2).Value
    
    sub_folder = output_folder & "\" & folder_name
    
    If fso.FolderExists(sub_folder) Then
        
        answer = MsgBox("The folder        '" & sub_folder & "' already exists." & vbNewLine & vbNewLine & _
                 "Do you want To continue running this vba code?", vbExclamation + vbYesNo, "Master Image Renaming")
        If answer = vbNo Then
            copy_rename_files = True
            Exit Function
        End If
    Else
        fso.CreateFolder output_folder & "\" & folder_name
    End If
    
    Call insert_log_row("Create " & folder_name & " & subfolder in the output folder.", "Success", size_name)
    
    On Error Resume Next
    For i = 1 To UBound(data_array, 1)
        input_file = data_array(i, 1)
        output_file = data_array(i, 2)
        
        input_path = image_folder & "\" & input_file
        output_path = sub_folder & "\" & output_file
        
        If output_file <> "" Then
            fso.CopyFile input_path, output_path, True
            error_message = Err.Description
            
            If Err.Number = 0 Then
                Call insert_log_row("Copied & renamed file.", "Success", size_name, input_file, output_file)
            Else
                Call insert_log_row("Copied & renamed file.", "Error", size_name, input_file, output_file, error_message)
            End If
            
            Err.Clear
        Else
            Call insert_log_row("Skipped copying file due To skip flag.", "Info", size_name, input_file, output_file)
        End If
    Next i
End Function

'The below subroutine is executed through the excel application and calls the other functions/subroutines as required.

Sub process_images()
    
    Dim fso         As New FileSystemObject
    Dim sh_rename   As Worksheet
    Dim sh_controls As Worksheet
    Dim sh_log      As Worksheet
    Dim answer      As Integer
    Dim timestamp   As Date
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sh_rename = rename
    Set sh_controls = controls
    Set sh_log = log
    
    'The below section of code launches a pop up dialogue box asking users if they are happy to proceed with the subroutine. If the user selects no the
    'the subroutine ends, if the user selects yes then the code will continue to run.
    
    answer = MsgBox("Are you ready To run the renaming code?" & vbNewLine & vbNewLine & _
             "Please back up your images & check the data before proceeding." & vbNewLine _
           & vbNewLine & "Click yes To run the code, no To exit." _
             , vbQuestion + vbYesNo + vbDefaultButton2, "Master Image Renaming")
    
    If answer = vbNo Then
        
        Exit Sub
        
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    timestamp = Now
    
    sh_log.Range("Start_Timestamp").Value = timestamp
    sh_log.Range("Folder_Image").Value = sh_controls.Range("Image_Folder").Value
    sh_log.Range("Folder_Output").Value = sh_controls.Range("Output_Folder").Value
    sh_log.Range("File_Line_List").Value = sh_controls.Range("Output_Folder").Value
    
    If sh_controls.Range("Customers_Check") = True Then
        
        If copy_rename_files(fso, "Customers") Then
            
            Exit Sub
            
        End If
        
    End If
    
    If sh_controls.Range("D2C_Check") = True Then
        
        If copy_rename_files(fso, "D2C") Then
            
            Exit Sub
            
        End If
        
    End If
    
    If sh_controls.Range("Size_1_Check") = True Then
        
        If copy_rename_files(fso, sh_controls.Range("Size_1_Check").Offset(0, -1).Value) Then
            
            Exit Sub
            
        End If
        
    End If
    
    If sh_controls.Range("Size_2_Check") = True Then
        
        If copy_rename_files(fso, sh_controls.Range("Size_2_Check").Offset(0, -1).Value) Then
            
            Exit Sub
            
        End If
        
    End If
    
    If sh_controls.Range("Size_3_Check") = True Then
        
        If copy_rename_files(fso, sh_controls.Range("Size_3_Check").Offset(0, -1).Value) Then
            
            Exit Sub
            
        End If
        
    End If
    
    If sh_controls.Range("Size_4_Check") = True Then
        
        If copy_rename_files(fso, sh_controls.Range("Size_4_Check").Offset(0, -1).Value) Then
            
            Exit Sub
            
        End If
        
    End If
    
    If sh_controls.Range("Size_5_Check") = True Then
        
        If copy_rename_files(fso, sh_controls.Range("Size_5_Check").Offset(0, -1).Value) Then
            
            Exit Sub
            
        End If
        
    End If
    
    If sh_controls.Range("Size_6_Check") = True Then
        
        If copy_rename_files(fso, sh_controls.Range("Size_6_Check").Offset(0, -1).Value) Then
            
            Exit Sub
            
        End If
        
    End If
    
    If sh_controls.Range("Size_7_Check") = True Then
        
        If copy_rename_files(fso, sh_controls.Range("Size_7_Check").Offset(0, -1).Value) Then
            
            Exit Sub
            
        End If
        
    End If
    
    sh_rename.Range("Rename_Selection").Value = "Customers"
    
    Call write_logs_to_table
    
    timestamp = Now
    
    sh_log.Range("End_Timestamp").Value = timestamp
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Dim Msg, Style, Title, Box
    
    Msg = "The image renaming Is complete & the New images can be found at:" & vbNewLine & vbNewLine & _
          sh_controls.Range("Output_Folder").Value & vbNewLine & vbNewLine & "On: " & CStr(timestamp)
    Style = vbOKOnly + vbInformation
    Title = "Master Image Renaming"
    Box = MsgBox(Msg, Style, Title)
    
End Sub

