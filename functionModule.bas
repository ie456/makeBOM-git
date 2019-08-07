Attribute VB_Name = "functionModule"
Function unUsedSheet(porTectAry As Variant)
    'Protect sheets

sheetProtect = porTectAry


'REMOVE UNUSED SHEETS
 With Workbooks(get_FileName_MakeBOM)

        For Each Worksheet In .Worksheets
            aa = Worksheet.Name
            If functionModule.IsInArray(Worksheet.Name, sheetProtect) = False Then
                Application.DisplayAlerts = False
                .Sheets(aa).Delete
                Application.DisplayAlerts = True
            End If
        Next

 End With

End Function

Function IsInArray(stringToBeFound As Variant, arr As Variant) As Boolean
  IsInArray = UBound(filter(arr, stringToBeFound)) > -1
End Function

Function IsLike(data As Variant, likeSample As Variant) As Boolean
    
    temp = Split(likeSample, "/")
    
    For i = 0 To UBound(temp)
    
        If data Like temp(i) Then
        
            IsLike = True
            
            GoTo exit1:
        
        End If
    
    Next
    

    IsLike = False
exit1:
    
    
End Function

Function protectSheet(SheetName As Variant)
With Workbooks(get_FileName_MakeBOM)
    .Worksheets(SheetName).Protect "123"
End With
End Function


Function unProtectSheet(SheetName As Variant)
With Workbooks(get_FileName_MakeBOM)
    .Worksheets(SheetName).Unprotect "123"
End With
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function updateUserFormValue(caseValue As Integer)

    
    Select Case caseValue
    
    Case 0
        updateUserFormValue_Page1
        updateUserFormValue_Page2
         updateUserFormValue_Page3
    Case 1
        updateUserFormValue_Page1
    Case 2
        updateUserFormValue_ComoList1
        'updateUserFormValue_Page2_excelSheet
    Case 22
        updateUserFormValue_ComoList2
        'updateUserFormValue_Page2_excelSheet2
    Case 23
        updateUserFormValue_ComoList3
        'updateUserFormValue_Page2_excelSheet2
    Case 3
        updateUserFormValue_checklv
    Case 4
        updateUserFormValue_Page3
        
    Case 5
        updateUserFormValue_SetNewValue
    Case 6
        updateUserFormValue_SetNewValue_row
    Case Else
         
    End Select
    
    

End Function


Private Function updateUserFormValue_Page3()

With Workbooks(get_FileName_MakeBOM)

    UserForm1.TextBox2.Visible = True
    UserForm1.TextBox3.Visible = True
    UserForm1.Label3.Visible = True
    UserForm1.Label4.Visible = True
    UserForm1.TextBox1 = .Worksheets("MAIN").Range("B24").Value
    UserForm1.TextBox2 = .Worksheets("MAIN").Range("B25").Value
    UserForm1.TextBox3 = .Worksheets("MAIN").Range("B26").Value
    UserForm1.TextBox4 = .Worksheets("MAIN").Range("B27").Value
    
    Select Case .Worksheets("MAIN").Range("J22").Value
    
    Case 3
        UserForm1.TextBox2.Visible = False
        UserForm1.TextBox3.Visible = False
        UserForm1.Label3.Visible = False
        UserForm1.Label4.Visible = False
    Case 4
        UserForm1.TextBox3.Visible = False
        UserForm1.Label4.Visible = False
    Case Else
    
    End Select
    
    
   
End With
End Function

Private Function updateUserFormValue_Page1()

With Workbooks(get_FileName_MakeBOM)
    UserForm1.TextBox5 = .Worksheets("MAIN").Range("B29").Value
    updateUserFormValue_ComoList2
    UserForm1.TextBox7 = .Worksheets("MAIN").Range("B31").Value
    updateUserFormValue_ComoList1
End With
End Function

Private Function updateUserFormValue_ComoList1()
     
 With Workbooks(get_FileName_MakeBOM)
     
     UserForm1.ComboBox1.Clear
     tempList = Split(.Worksheets("MAIN").Range("B32"), ",")
     
     For Each tempItem In tempList
     
        UserForm1.ComboBox1.AddItem tempItem
     
     Next
     
     On Error GoTo exit1
     UserForm1.ComboBox1.Text = UserForm1.ComboBox1.List(0)
End With
exit1:
     
End Function
Private Function updateUserFormValue_ComoList2()
     
 With Workbooks(get_FileName_MakeBOM)
     
     UserForm1.ComboBox2.Clear
     tempList = Split(.Worksheets("MAIN").Range("B30"), ",")
     
     For Each tempItem In tempList
     
        UserForm1.ComboBox2.AddItem tempItem
     
     Next
     
     On Error GoTo exit1
     UserForm1.ComboBox2.Text = UserForm1.ComboBox2.List(0)
End With
exit1:
     
End Function
Private Function updateUserFormValue_ComoList3()
     
 With Workbooks(get_FileName_MakeBOM)
     
     UserForm1.ComboBox3.Clear
     tempList = Split(.Worksheets("MAIN").Range("E30"), ",")
     
     For Each tempItem In tempList
     
        UserForm1.ComboBox3.AddItem tempItem
     
     Next
     
     On Error GoTo exit1
     UserForm1.ComboBox3.Text = UserForm1.ComboBox3.List(0)
End With
exit1:
     
End Function

Private Function updateUserFormValue_Page2_excelSheet()

With Workbooks(get_FileName_MakeBOM)
     
     UserForm1.ComboBox1.Clear
     tempList = Split(.Worksheets("MAIN").Range("B32"), ",")
     
     For Each tempItem In tempList
     
        UserForm1.ComboBox1.AddItem tempItem
     
     Next
     
     On Error GoTo exit1
     UserForm1.ComboBox1.Text = UserForm1.ComboBox1.List(0)
End With
exit1:
End Function
Private Function updateUserFormValue_Page2_excelSheet2()

With Workbooks(get_FileName_MakeBOM)
     
     UserForm1.ComboBox2.Clear
     tempList = Split(.Worksheets("MAIN").Range("B30"), ",")
     
     For Each tempItem In tempList
     
        UserForm1.ComboBox2.AddItem tempItem
     
     Next
     
     On Error GoTo exit1
     UserForm1.ComboBox2.Text = UserForm1.ComboBox2.List(0)
End With
exit1:
End Function

Private Function updateUserFormValue_Page2()
    With Workbooks(get_FileName_MakeBOM)
        UserForm1.TextBox8.Value = .Worksheets("MAIN").Range("B34").Value
        UserForm1.TextBox9.Value = .Worksheets("MAIN").Range("B35").Value
        UserForm1.TextBox10.Value = .Worksheets("MAIN").Range("B36").Value
        UserForm1.TextBox11.Value = .Worksheets("MAIN").Range("B37").Value
        UserForm1.TextBox14.Value = .Worksheets("MAIN").Range("B38").Value
        UserForm1.TextBox16.Value = .Worksheets("MAIN").Range("B39").Value
        UserForm1.TextBox15.Value = .Worksheets("MAIN").Range("B40").Value
        
        UserForm1.TextBox20.Value = .Worksheets("MAIN").Range("B34").Value
        UserForm1.TextBox21.Value = .Worksheets("MAIN").Range("B35").Value
        UserForm1.TextBox22.Value = .Worksheets("MAIN").Range("B36").Value
        UserForm1.TextBox24.Value = .Worksheets("MAIN").Range("B37").Value
        UserForm1.TextBox23.Value = .Worksheets("MAIN").Range("B38").Value
        UserForm1.TextBox25.Value = .Worksheets("MAIN").Range("B39").Value
        UserForm1.TextBox19.Value = .Worksheets("MAIN").Range("B40").Value
    
End With
End Function
Private Function updateUserFormValue_checklv()
    
    If UserForm1.TextBox2.Value = "Please Enter lv4 PartNumber" Or Replace(UserForm1.TextBox2.Value, " ", "") = "" Then
        UserForm1.TextBox5.Enabled = False
        'UserForm1.TextBox6.Enabled = False
        
        UserForm1.CommandButton6.Enabled = False
        'UserForm1.CommandButton7.Enabled = False
        
        GoTo exit1
    Else
        UserForm1.TextBox5.Enabled = True
        'UserForm1.TextBox6.Enabled = True
        
        UserForm1.CommandButton6.Enabled = True
        'UserForm1.CommandButton7.Enabled = True
    End If
    
    If UserForm1.TextBox3.Value = "Please Enter lv5 PartNumber" Or Replace(UserForm1.TextBox3.Value, " ", "") = "" Then
        
        'UserForm1.TextBox6.Enabled = False
        'UserForm1.CommandButton7.Enabled = False
        
    Else
        'UserForm1.TextBox6.Enabled = True
        'UserForm1.CommandButton7.Enabled = True
    End If
    
    
exit1:
    
End Function
Private Function updateUserFormValue_SetNewValue()
     'SETTING PAGE NEW LOAD EXCEL FILE
With Workbooks(get_FileName_MakeBOM)
    UserForm1.TextBox8.Value = .Worksheets("MAIN").Range("B34").Value
    UserForm1.TextBox9.Value = .Worksheets("MAIN").Range("B35").Value
    UserForm1.TextBox10.Value = .Worksheets("MAIN").Range("B36").Value
    UserForm1.TextBox11.Value = .Worksheets("MAIN").Range("B37").Value
    UserForm1.TextBox14.Value = .Worksheets("MAIN").Range("B38").Value
    UserForm1.TextBox16.Value = .Worksheets("MAIN").Range("B39").Value
    
    UserForm1.TextBox20.Value = .Worksheets("MAIN").Range("B34").Value
    UserForm1.TextBox21.Value = .Worksheets("MAIN").Range("B35").Value
    UserForm1.TextBox22.Value = .Worksheets("MAIN").Range("B36").Value
    UserForm1.TextBox24.Value = .Worksheets("MAIN").Range("B37").Value
    UserForm1.TextBox23.Value = .Worksheets("MAIN").Range("B38").Value
    UserForm1.TextBox25.Value = .Worksheets("MAIN").Range("B39").Value
    
    
    
End With
End Function

Private Function updateUserFormValue_SetNewValue_row()
     'SETTING PAGE NEW LOAD EXCEL FILE
With Workbooks(get_FileName_MakeBOM)
     
    UserForm1.TextBox15.Value = .Worksheets("MAIN").Range("B40").Value
         
    UserForm1.TextBox19.Value = .Worksheets("MAIN").Range("B40").Value
End With
End Function





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function getPath(SheetName As String, saveLocaPath As String, saveLocaSheet As String, upSelNum As Integer) As String


    Dim isExcelFile As Boolean
    'Dim excelFileType
    
    'excelFileType = Array("xls", "xlsx", "xlsm")
    
    'tempValue = getPath_Ex
    tempValue = getPath_ExRptBom
    tempAry = Split(tempValue, "\")
    
    
    
    With Workbooks(get_FileName_MakeBOM)
    
    
        If tempValue = "" Then
            getPath = .Worksheets(SheetName).Range(saveLocaPath).Value
        Else
                
                Application.ScreenUpdating = False
                functionModule.unProtectSheet (SheetName)
                .Worksheets(SheetName).Range(saveLocaPath).Value = tempValue
                
                
                
                
                'Get worksheets list
               
                temp = ""
                
                If checkFileType(tempValue) = 0 Then
                    Workbooks.Open FileName:=tempValue
                    
                    
                    For Each tempSheet In Worksheets
                    'ComboBox1.AddItem tempSheet.Name
                        If tempSheet.Name = Worksheets.item(1).Name Then
                            temp = tempSheet.Name
                        Else
                            temp = temp & "," & tempSheet.Name
                        End If
                    Next
                    
                     Workbooks(tempAry(UBound(tempAry))).Close SaveChanges:=False
                    
                End If
                
                
                
               
                'ComboBox1.Text = ComboBox1.List(0)
        
                'tempAry = Split(tempValue, "\")
               
                
                .Worksheets(SheetName).Range(saveLocaSheet).Value = temp
                functionModule.protectSheet (SheetName)
                Application.ScreenUpdating = True
                getPath = tempValue
                updateUserFormValue (upSelNum)
            
        End If
    
    
 
    End With

    
    

End Function
'Select Bom file
'*.rpt>> Concepte
'*.xls;*.xlsx;*.xlsm >> Excel file
'*.Bom;>> Orcad file
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function getPath_Gen(SheetName As String, saveLocaPath As String, saveLocaSheet As String, upSelNum As Integer, filterString As Variant) As String


    Dim isExcelFile As Boolean
    'Dim excelFileType
    
    'excelFileType = Array("xls", "xlsx", "xlsm")
    
    'tempValue = getPath_Ex
    tempValue = getPath_Ex(filterString)
    tempAry = Split(tempValue, "\")
    
    
    
    With Workbooks(get_FileName_MakeBOM)
    
    
        If tempValue = "" Then
            getPath_Gen = .Worksheets(SheetName).Range(saveLocaPath).Value
        Else
                
                Application.ScreenUpdating = False
                functionModule.unProtectSheet (SheetName)
                .Worksheets(SheetName).Range(saveLocaPath).Value = tempValue
                
                
                
                
                'Get worksheets list
               
                temp = ""
                
                If checkFileType(tempValue) = 0 Then
                    Workbooks.Open FileName:=tempValue
                    
                    
                    For Each tempSheet In Worksheets
                    'ComboBox1.AddItem tempSheet.Name
                        If tempSheet.Name = Worksheets.item(1).Name Then
                            temp = tempSheet.Name
                        Else
                            temp = temp & "," & tempSheet.Name
                        End If
                    Next
                    
                     Workbooks(tempAry(UBound(tempAry))).Close SaveChanges:=False
                    
                End If
                
                
                
               
                'ComboBox1.Text = ComboBox1.List(0)
        
                'tempAry = Split(tempValue, "\")
               
                
                .Worksheets(SheetName).Range(saveLocaSheet).Value = temp
                functionModule.protectSheet (SheetName)
                Application.ScreenUpdating = True
                getPath_Gen = tempValue
                updateUserFormValue (upSelNum)
            
        End If
    
    
 
    End With

    
    

End Function
'Select Bom file
'*.rpt>> Concepte
'*.xls;*.xlsx;*.xlsm >> Excel file
'*.Bom;>> Orcad file


Function getPath_ExRptBom() As String

        'Micro_book = Application.ActiveWorkbook.Name
        Micro_book = get_FileName_MakeBOM
    

    Dim filter, caption, datafilename, cmpsheet As String
    
    
    filter = "BOM file (*.rpt;*.xls;*.xlsx;*.xlsm;), *.rpt;*.xls;*.xlsx;*.xlsm"
    caption = "Select a NET file"
    datafilename = Application.GetOpenFilename(filter, , caption)

    'If datafilename = fail Then Exit Sub  'Do nothing when dont select file
    
    
    If datafilename = False Then
        getPath_ExRptBom = ""
    Else
        getPath_ExRptBom = datafilename
    End If
    
    
    
End Function
Function getPath_Ex(filterString As Variant) As String

        'Micro_book = Application.ActiveWorkbook.Name
        Micro_book = get_FileName_MakeBOM
    

    Dim filter, caption, datafilename, cmpsheet As String
    
    
    filter = filterString ' "Component file (*.xls;*.xlsx;*.csv), *.xls;*.xlsx;*.csv"
    caption = "Select a file"
    datafilename = Application.GetOpenFilename(filter, , caption)

    'If datafilename = fail Then Exit Sub  'Do nothing when dont select file
    
    
    If datafilename = False Then
        getPath_Ex = ""
    Else
        getPath_Ex = datafilename
    End If
    
    
    
End Function

Function getPath_Txt(txtString As String) As String
    'Open nets file

    Micro_book = get_FileName_MakeBOM

    Dim filter, caption, datafilename, cmpsheet As String

    tmepString = txtString


    filter = "txt file (*.txt), *.txt"
    caption = "Select a NET file"
    datafilename = Application.GetOpenFilename(filter, , caption)

    'If datafilename = fail Then Exit Sub  'Do nothing when dont select file
    
    
    If datafilename = False Then
        getPath_Txt = txtString
    Else
        getPath_Txt = datafilename
        
    End If
    
    

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'0 -> EXCEL FILE (.xls ; .xlsx ; .xlsm)
'1 -> CONCEPT FILE (.rpt)
'2 -> Orcad FILE (.BOM)

Function checkFileType(checkFileName As Variant) As Integer

    fileNamExt = Split(checkFileName, ".")
    
    
    Select Case LCase(fileNamExt(UBound(fileNamExt)))
    
    Case "xls", "xlsx", "xlsm", "csv"
        checkFileType = 0
    Case "rpt"
        checkFileType = 1
    Case "bom"
        checkFileType = 2
    Case Else
        checkFileType = 100
    End Select

End Function



Function getStr(tempString As Variant) As String

    Dim tempStr As String

    For i = 1 To Len(tempString)
    
        Select Case Asc(Mid(tempString, i, 1))
        Case 65 To 90
            tempStr = tempStr & Mid(tempString, i, 1)
        End Select
    
    Next
    
    getStr = tempStr
    
End Function


Function check_345Vail() As Boolean

End Function


Function check_path() As Boolean

    
End Function

Function check_TopBotOnly() As Boolean

    
End Function



Function compareString(ByRef String1 As Variant, ByRef String2 As Variant, ignorType As Variant) As String
  
  Dim tempList1, tempList2
  Dim tempString
  Dim firstCount
  
  
  firstCount = 0
  
  
  tempAry = Split(String1, ignorType)
  Set tempList1 = aryToList(tempAry)
  
  tempAry = Split(String2, ignorType)
  Set tempList2 = aryToList(tempAry)
  
  Set tempList1Clone = tempList1.Clone
  
  
  
  For Each tempValue In tempList1
  
    If tempList2.contains(tempValue) Then
    
        tempList2.removeat (tempList2.indexof(tempValue, 0))
        tempList1Clone.removeat (tempList1Clone.indexof(tempValue, 0))
        
        If firstCount = 0 Then
            tempString = tempValue
            firstCount = 1
        Else
            tempString = tempString & "," & tempValue
        End If
        
        
    End If
  
  Next
  
  String1 = aryToString(tempList1Clone)
  
  String2 = aryToString(tempList2)
  
  compareString = tempString
End Function

Private Function aryToList(arr As Variant) As Object
    
    Set aryToList = CreateObject("system.collections.arraylist")
    
    For Each tempValue In arr
        aryToList.Add tempValue
    Next
    
End Function

Function aryToString(arr As Variant) As String
    temp = ""
    
    For Each tempValue In arr
        temp = temp & tempValue & ","
    Next
    
    If temp <> "" Then temp = Left(temp, Len(temp) - 1)
    
    aryToString = temp
    
    
End Function



Function printDataInSheet(startRow As Integer, SheetName As Variant, parent As Variant, _
                        arrlist1 As Variant, ar1_indexcol1 As Variant, _
                        arrlist2 As Variant, ar1_indexcol2 As Variant, _
                        arrlist3 As Variant, ar1_indexcol3 As Variant) As Integer
                        
                        
                        
    index = startRow
    
    
  
    With Workbooks(get_FileName_MakeBOM).Worksheets(SheetName)
    
        For i = 0 To arrlist1.Count - 1
        
            .Cells(index + i, 1) = parent
            .Cells(index + i, 3) = (i + 1) * 10
            .Cells(index + i, ar1_indexcol1) = arrlist1(i)
            .Cells(index + i, ar1_indexcol2) = arrlist2(i)
            .Cells(index + i, ar1_indexcol3) = arrlist3(i)
            
        Next i
    
    End With
    
    
    printDataInSheet = index + i

End Function
Function delSheet(SheetName As Variant, colorOnOff As Boolean, colorCode As Integer)
    
    temp = 0
    
    'check sheet name
    On Error GoTo exit1
    
    With Workbooks(tempWorkBookName)
    
                For Each tempSheet In .Worksheets
                    
                    If tempSheet.Name = SheetName Then
                    
                        .Worksheets(SheetName).Delete
                        GoTo exit1
                    
                    End If
                
                Next
                
                
                    
            
            
            temp = 1
            
exit1:
                    
                    If colorOnOff Then
                            .Worksheets(SheetName).Tab.ColorIndex = colorCode
                    End If
                    
                delSheet = temp
        
    End With

End Function


Function checkWord(item As Variant) As Boolean


    If (64 < Asc(item.Value) And Asc(item.Value) < 91) Then
        checkWord = True
    Else
        checkWord = False
    End If

End Function


Function printOut_dir(filname As Variant, Path As Variant) As String

    Dim tempListAry
    
     With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    
    timeout = 0
    tempCountName = ""
    tempfileName = filname
   Do While 1
    
    
     If Dir(Path & "\" & tempfileName & tempCountName & ".xls", vbDirectory) = vbNullString Then
        
        If tempCountName <> "" Then
            tempfileName = tempfileName & tempCountName
        End If
        
        
        Exit Do
     Else
        tempCountName = "_" & timeout
     End If
    
    
    
        If timeout > 50 Then Exit Do
            timeout = timeout + 1
   
   Loop
    
    
    With Workbooks(get_FileName_MakeBOM)
    
    
    
    
    Set wb2 = Application.Workbooks.Add
    
    SheetName = Workbooks(get_FileName_MakeBOM).Worksheets("MAIN").Range("B24").Value
    .Worksheets(SheetName).Copy after:=wb2.Sheets(wb2.Sheets.Count)
   
   
    wb2.Sheets(1).Delete
    
   
    wb2.SaveAs FileName:=Path & "\" & tempfileName & ".xls", FileFormat:=56
    'Wb2.SaveAs fileName:=wb1.Path & "\" & filname & ".xlsx", FileFormat:=51
    wb2.Close
 End With
 
 

printOut_dir = tempfileName
 
'MsgBox "Done!"
    

'EVENT no sheet select


exitEnd:


With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
End With

    Exit Function






End Function
