Option Explicit

'===================================================================================================================
'Project: Master Image Renaming
'Language: VBA
'Author: Leo Pickard
'Date: March 2025
'Version: 2.0

'This vba file contains subroutines and functions which are not used in the image renaming process. The below code
'adds fucntionality for picking files/folders & refreshing data.
'===================================================================================================================

Function pick_folder(Optional dialogTitle As String = "Select a Folder", Optional button_desc As String = "Select")
    Dim objFileDialog As FileDialog
    Dim objSelectedFolder As Variant
    
    Set objFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With objFileDialog
        .ButtonName = button_desc
        .Title = dialogTitle
        .InitialView = msoFileDialogViewList
        
        If .Show = -1 Then
            pick_folder = .SelectedItems(1)
        Else
            pick_folder = ""
        End If
    End With
    
    Set objFileDialog = Nothing
End Function

Function pick_excel_file(Optional dialogTitle As String = "Select an Excel Line List", Optional button_desc As String = "Select")
    Dim objFileDialog As FileDialog
    Dim objSelectedFile As Variant
    
    Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With objFileDialog
        .ButtonName = button_desc
        .Title = dialogTitle
        .InitialView = msoFileDialogViewList
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            pick_excel_file = .SelectedItems(1)
        Else
            pick_excel_file = ""
        End If
    End With
    
    Set objFileDialog = Nothing
End Function

Sub set_image_folder_path()
    
    Dim image_folder_path As String
    Dim ws_controls As Worksheet
    
    Set ws_controls = controls
    
    image_folder_path = pick_folder("Select Source Image Folder", "Select Image Folder")
    
    If image_folder_path <> "" Then
        ws_controls.Range("Image_Folder").Value = image_folder_path
    End If
    
End Sub

Sub set_output_folder_path()
    
    Dim output_folder_path As String
    Dim ws_controls As Worksheet
    
    Set ws_controls = controls
    
    output_folder_path = pick_folder("Select Output Folder", "Select Output Folder")
    
    If output_folder_path <> "" Then
        ws_controls.Range("Output_Folder").Value = output_folder_path
    End If
    
End Sub

Sub set_line_list_file_path()
    
    Dim file_path   As String
    Dim ws_controls As Worksheet
    
    Set ws_controls = controls
    
    file_path = pick_excel_file("Select Excel Line List File", "Select Line List")
    
    If file_path <> "" Then
        
        ws_controls.Range("Line_List_File").Value = file_path
        
    End If
    
End Sub

Sub refresh_line_list()
    
    Dim answer      As Integer
    Dim time
    
    answer = MsgBox("Are you sure you want To refresh the line list data?" & vbNewLine & vbNewLine & _
             "Please note that any manual changes you have made To the line list worksheet will be overwritten." & _
             vbNewLine & vbNewLine & "Click yes To refresh the line list data, no To exit." _
             , vbQuestion + vbYesNo + vbDefaultButton2, "Master Image Renaming")
    
    If answer = vbNo Then
        
        Exit Sub
        
    End If
    
    ThisWorkbook.Connections("Query - Line List").OLEDBConnection.Refresh
    
    time = CStr(Now)
    
    MsgBox "The line list data has been successfully refreshed:" & vbNewLine & vbNewLine & "@ " & time, vbInformation + vbOKOnly, "Master Image Renaming"
    
End Sub

Sub refresh_renaming_data()
    
    Dim answer      As Integer
    Dim time
    
    answer = MsgBox("Are you sure you want To refresh the final renaming data?" & vbNewLine & vbNewLine & _
             "Please note that any manual changes you have made To the renaming data worksheet will be overwritten." & _
             vbNewLine & vbNewLine & "Click yes To refresh the line list data, no To exit." _
             , vbQuestion + vbYesNo + vbDefaultButton2, "Master Image Renaming")
    
    If answer = vbNo Then
        
        Exit Sub
        
    End If
    
    ThisWorkbook.Connections("Query - Renaming Data(1)").OLEDBConnection.Refresh
    
    'ThisWorkbook.Connections("Query - Item Data").OLEDBConnection.Refresh
    
    time = CStr(Now)
    
    MsgBox "The image renaming data has been successfully refreshed:" & vbNewLine & vbNewLine & "@ " & time, vbInformation + vbOKOnly, "Master Image Renaming"
    
End Sub
