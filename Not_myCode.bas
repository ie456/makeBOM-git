Attribute VB_Name = "Not_myCode"
'
'
'
'Thanks for Chih-Da Wu
'
'
'



Function GenerateComponentCSVFile(ByVal strSourceFile As String, ByVal strDestinationFile As String) As Integer

   Dim i As Integer
   Dim fs, fd As Integer
   
   Dim line_buf As String
   Dim data_buf() As String
   Dim current_row As Integer
   
   
   Dim refdes As String
   Dim dev_type As String
   Dim comp_value As String
   Dim comp_tol As String
   Dim comp_package As String
   Dim sym_x As String
   Dim sym_y As String
   Dim sym_rotate As String
   Dim sym_mirror As String
   Dim sym_buf() As String
   
   Dim d As Date
   Dim item_complete As Boolean
   Dim item_imported As Integer
   Dim item_expected As Integer
   

   item_imported = 0
   item_expected = 0
   
   fs = FreeFile
   Open strSourceFile For Input As #fs

   fd = FreeFile
   Open strDestinationFile For Output As #fd
   
      ' ----------- Get Total item number -----------------
   Do
      line_buf = getLine_UnixDos(fs)
   Loop Until ((InStr(1, line_buf, "LISTING", vbTextCompare)) Or EOF(fs))
   data_buf = Split(line_buf, " ")
   
    ' ------------ Output header to destination -----------------
   Print #fd, "Design Name: " & strSourceFile & ",,,,,,,,"
   
   item_expected = CInt(Trim(data_buf(1)))
   
   d = Now
   Print #fd, "Date: " & FormatDateTime(d, vbLongDate) & " " & FormatDateTime(d, vbLongTime) & ",,,,,,,,"
   Print #fd, "Total Components: " & Trim(data_buf(1)) & ",,,,,,,,"
   Print #fd, ",,,,,,,,"
   Print #fd, "Component Report" & ",,,,,,,,"
   Print #fd, "REFDES,COMP_DEVICE_TYPE,COMP_VALUE,COMP_TOL,COMP_PACKAGE,SYM_X,SYM_Y,SYM_ROTATE,SYM_MIRROR"
   
   
    ' ------------ Get Components Properties -----------------
 
   Do
        line_buf = getLine_UnixDos(fs)

       'New Item start
     If (InStr(1, line_buf, "Item", vbBinaryCompare)) Then
        'New Item start, Clear all property value
         refdes = ""
         dev_type = ""
         comp_value = ""
         comp_tol = ""
         comp_package = ""
         sym_x = ""
         sym_y = ""
         sym_rotate = ""
         sym_mirror = ""
         
         
         item_complete = False
         
     Else
         
         If item_complete = False Then
            data_buf = Split(line_buf, ":")
            If UBound(data_buf) >= 0 Then
            Select Case Trim(data_buf(0))
                Case "Reference Designator"
                    refdes = Trim(data_buf(1))
                Case "Package Symbol"
                    comp_package = Trim(data_buf(1))
                Case "Device Type"
                    dev_type = Chr(34) & Trim(data_buf(1)) & Chr(34)      ' Add ""
                Case "Value"
                    comp_value = Trim(data_buf(1))
                Case "Tolerance"
                    comp_tol = Trim(data_buf(1))
                Case "origin-xy"
                    sym_buf = Split(Replace(Replace(Trim(data_buf(1)), "(", ""), ")", ""), " ")
                    sym_x = Trim(sym_buf(0))
                    sym_y = Trim(sym_buf(1))
                Case "rotation"
                    sym_buf = Split(Trim(data_buf(1)), " ")
                    sym_rotate = sym_buf(0)
                Case "not_mirrored"
                   sym_mirror = "NO"
                   item_complete = True
                Case "mirrored"
                   sym_mirror = "YES"
                   item_complete = True
              End Select
            
              If item_complete = True Then
                   'Output result to destination file
                   Print #fd, refdes & "," & dev_type & "," & comp_value & "," & comp_tol & "," & comp_package & "," & sym_x & "," & sym_y & "," & sym_rotate & "," & sym_mirror
                   item_imported = item_imported + 1
              End If
            End If
         End If
 
       
    End If
       
     
   Loop Until EOF(fs)

   Close #fs
   Close #fd
   
   If item_imported = item_expected Then
       GenerateComponentCSVFile = item_imported
   Else
       GenerateComponentCSVFile = -1
   End If
End Function

Sub btnGenerateComponentCSV_Click()
    Dim strSourceFile As String
    Dim strDestinationFile As String
    Dim i As Integer
    
    strSourceFile = SelectFile()
    strDestinationFile = SaveAsFile()
    
    If strSourceFile = "" Or strDestinationFile = "" Then
        Exit Sub
    Else
        i = GenerateComponentCSVFile(strSourceFile, strDestinationFile)
        If i >= 0 Then
            MsgBox ("Success." & CStr(i) & " items tranferred.")
            'ThisWorkbook.Sheets("Control").Cells(6, 2) = strDestinationFile
            Call subModule.saveData("MAIN", strDestinationFile, "B29")
            tempStr = Split(strDestinationFile, "\")
            Call subModule.saveData("MAIN", Replace(tempStr(UBound(tempStr)), ".csv", ""), "B30")
            functionModule.updateUserFormValue (1)
            
        Else
            MsgBox ("Some problem!! Quantity mismatch while transfer")
        End If
    End If
    
    
End Sub


Function SelectFile() As String

    Dim myFd As FileDialog
    Set myFd = Application.FileDialog(msoFileDialogFilePicker)
        With myFd
        .Title = "Please select the source file"
        .ButtonName = "OK"
        .InitialFileName = ThisWorkbook.Path & "\"
        If .show = True Then
            SelectFile = .SelectedItems(1)
        Else
            MsgBox "Cancelled"
'            SelectFile = ""
            Exit Function
        End If
    End With

End Function

Function SaveAsFile() As String

    Dim myFd As FileDialog
    Dim a As Integer
    
    Set myFd = Application.FileDialog(msoFileDialogSaveAs)
        With myFd
        .Title = "Please select the destination file"
        .ButtonName = "OK"
        .InitialFileName = ThisWorkbook.Path & "\"
        .AllowMultiSelect = False
        
        For a = 1 To .Filters.Count
            If (LCase(.Filters(a).Extensions) = "*.csv") Then

                .FilterIndex = a
            End If
        Next a
        
        If .show = True Then
            SaveAsFile = .SelectedItems(1)
        Else
            MsgBox "Cancelled"
            SaveAsFile = ""
            Exit Function
        End If
    End With

End Function
Function getLine_UnixDos(ByVal filenum As Integer) As String

    Dim thischar As String
    Dim thisline As String
    Dim lastchar As String
    Dim linelen As Integer
    
        
    While Not EOF(1)
        
        thischar = Input(1, #filenum)
        If thischar = vbLf Then
            lastchar = Right(thisline, 1)
            If lastchar = vbCr Then
                linelen = Len(thisline)
                thisline = Left(thisline, linelen - 1)
            End If
            getLine_UnixDos = thisline
            Exit Function
        
        Else
            thisline = thisline & thischar
        End If
    Wend
     getLine_UnixDos = thisline
    'strLine = thisline
    
End Function
