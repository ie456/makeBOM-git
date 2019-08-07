Attribute VB_Name = "subModule"
Sub initSheet()

Dim protectSheet

     'If Workbooks(get_FileName_MakeBOM).Sheets("MAIN").Range("B22") = 0 Then
        
        protectSheet = Array("MAIN")

        functionModule.unUsedSheet (protectSheet)
        
        'End If
        
End Sub
'------------------------------------------------------
'---This function is uesed for fill in vale on cells---
'------------------------------------------------------
Sub saveData(SheetName As Variant, data As Variant, Location As Variant)

With Workbooks(get_FileName_MakeBOM)

    unProtectSheet (SheetName)
    
    .Worksheets(SheetName).Range(Location).Value = data

    protectSheet (SheetName)

End With

End Sub



 Sub printDataInSheet_txt(startRow As Integer, Index_column As Integer, SheetName As Variant, arrlist As Variant)
    
    index = startRow
    
    
  
    With Workbooks(get_FileName_MakeBOM).Worksheets(SheetName)
    
        For i = 0 To arrlist.Count - 1
        
            .Cells(index + i, Index_column) = arrlist(i)
            
        
        Next i
    
    End With
    
End Sub

 Sub printDataInSheet_lv(startRow As Integer, SheetName As Variant, _
                        arrlist1 As Variant, ar1_indexcol1 As Variant, _
                        arrlist2 As Variant, ar1_indexcol2 As Variant, _
                        arrlist3 As Variant, ar1_indexcol3 As Variant)
    
    
    index = startRow
    
    
  
    With Workbooks(get_FileName_MakeBOM).Worksheets(SheetName)
    
        For i = 0 To arrlist1.Count - 1
        
            .Cells(index + i, 1) = SheetName
            .Cells(index + i, 3) = (i + 1) * 10
            .Cells(index + i, ar1_indexcol1) = arrlist1(i)
            .Cells(index + i, ar1_indexcol2) = arrlist2(i)
            .Cells(index + i, ar1_indexcol3) = arrlist3(i)
            
        Next i
    
    End With
    
End Sub

 Sub printDataInSheet_check(startRow As Integer, SheetName As Variant, arrlist1 As Variant, arrlist2 As Variant, arrlist3 As Variant)
    
    
    
    Dim tempAry
    
    index = startRow
    
    
  
    With Workbooks(get_FileName_MakeBOM).Worksheets(SheetName)
    
        For i = 0 To arrlist1.Count - 1
            
            If arrlist2(i) <> "" Then
                tempAry = Split(arrlist2(i), ",")
            Else
                ReDim tempAry(0 To arrlist3(i) - 1)
            End If
            For j = 0 To UBound(tempAry)
            
                .Cells(index, 1).Value = arrlist1(i)
                .Cells(index, 2).Value = tempAry(j)
                
                index = index + 1
            Next
            
            
            
            
            
        Next i
    
    End With
    
End Sub


Sub creatSheet(SheetName As Variant, colorOnOff As Boolean, colorCode As Integer)
    
    'check sheet name
    On Error GoTo exit1
    
    With Workbooks(get_FileName_MakeBOM)
    
                For Each tempSheet In .Worksheets
                    
                    If tempSheet.Name = SheetName Then
                    
                        .Worksheets(SheetName).Cells.Clear
                        '.Worksheets(sheetName).Buttons.Delete
                        GoTo exit1
                    
                    End If
                
                Next
                
                
                    .Sheets.Add(after:=Worksheets(.Worksheets.Count)).Name = SheetName
                    'ActiveSheet.Name = sheetName
            
            
exit1:
                    
                    If colorOnOff Then
                            .Worksheets(SheetName).Tab.ColorIndex = colorCode
                    End If
        
    End With

End Sub

Sub creatSheet_lv(SheetName As Variant, colorOnOff As Boolean, colorCode As Integer)
    
    'check sheet name
    On Error GoTo exit1
    
    With Workbooks(get_FileName_MakeBOM)
    
                For Each tempSheet In .Worksheets
                    
                    If tempSheet.Name = SheetName Then
                    
                        .Worksheets(SheetName).Cells.Clear
                        '.Worksheets(sheetName).Buttons.Delete
                        GoTo exit1
                    
                    End If
                
                Next
                
                
                    .Sheets.Add(after:=Worksheets(.Worksheets.Count)).Name = SheetName
                    'ActiveSheet.Name = sheetName
                
                
               

                
            
exit1:
                    
                    
                        .Sheets(SheetName).Cells(1, 1) = "Parent"
                        .Sheets(SheetName).Cells(1, 2) = "Part Number"
                        .Sheets(SheetName).Cells(1, 3) = "Item Number"
                        .Sheets(SheetName).Cells(1, 4) = "Alt Grp"
                        .Sheets(SheetName).Cells(1, 5) = "Usage(%)"
                        .Sheets(SheetName).Cells(1, 6) = "Qty"
                        .Sheets(SheetName).Cells(1, 7) = "Location"
                    
                    
                    
                    If colorOnOff Then
                            .Worksheets(SheetName).Tab.ColorIndex = colorCode
                    End If
        
    End With

End Sub

Sub creatSheet_General(SheetName As Variant, colorOnOff As Boolean, colorCode As Integer, Header As Variant)
    
    'check sheet name
    On Error GoTo exit1
    
    With Workbooks(get_FileName_MakeBOM)
    
                For Each tempSheet In .Worksheets
                    
                    If tempSheet.Name = SheetName Then
                    
                        .Worksheets(SheetName).Cells.Clear
                        '.Worksheets(sheetName).Buttons.Delete
                        GoTo exit1
                    
                    End If
                
                Next
                
                
                    .Sheets.Add(after:=Worksheets(.Worksheets.Count)).Name = SheetName
                    'ActiveSheet.Name = sheetName
                    
                    
                    
            
            
exit1:
                    For i = 0 To UBound(Header)
                        .Sheets(SheetName).Cells(1, i + 1) = Header(i)
                    Next
                    
                    
                    If colorOnOff Then
                            .Worksheets(SheetName).Tab.ColorIndex = colorCode
                    End If
        
    End With

End Sub



Sub open_ExcelFile(FileName As Variant, SheetName As Variant, activeBook As Variant, sheetSel As Variant)
    
    
    
    On Error GoTo Exception1:
    Workbooks.Open FileName:=FileName
    openActive = ActiveWorkbook.Name
   
   If Sheets.Count = 1 Then
    Sheets.Add after:=Sheets(Worksheets.Count)
   End If
   
   
   
    Workbooks(openActive).Sheets(sheetSel).Move after:=Workbooks(activeBook).Sheets(Workbooks(activeBook).Worksheets.Count)
    
    Workbooks(activeBook).Sheets(Worksheets.Count).Name = SheetName
    Workbooks(activeBook).Sheets(Worksheets.Count).Tab.ColorIndex = 41
    
    tempAry = Split(FileName, "\")
    Workbooks(tempAry(UBound(tempAry))).Close SaveChanges:=False
    
    
    Exit Sub
    
Exception1:
    MsgBox ("Please check the path " & FileName)
    
    
End Sub

Sub open_ConceptFile(FileName As Variant, SheetName As Variant, activeBook As Variant)

    Workbooks.Open FileName:=FileName
    
    'Rows("1:6").EntireRow.Delete
    
    Range("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1)), TrailingMinusNumbers:=True
    Range("A1").Select
    
    Sheets(1).Move after:=Workbooks(activeBook).Sheets(Workbooks(activeBook).Worksheets.Count)
    Workbooks(activeBook).Sheets(Worksheets.Count).Name = SheetName
    Workbooks(activeBook).Sheets(Worksheets.Count).Tab.ColorIndex = 41
    
End Sub
Sub inputInList_overwrite(ByRef arylist1 As Variant, insertData As Variant, index As Variant)
    
    arylist1.removeat (index)
    arylist1.Insert index, insertData
    

End Sub


Sub inputInList_Loc(ByRef arylist1 As Variant, insertData As Variant, index As Variant)

    
    
    tempString = arylist1(index) & "," & insertData
    
    arylist1.removeat (index)
    arylist1.Insert index, tempString
    

End Sub



Sub inputInList_QTY(ByRef arylist1 As Variant, insertNum As Variant, index As Variant)

    
    tempQTY = arylist1(index) + insertNum
    
    arylist1.removeat (index)
    arylist1.Insert index, tempQTY
    

End Sub


'Sub inputInList(ByVal arylist_PN As Variant, insertData_PN As Variant, _
'                ByVal arylist_QTY As Variant, insertData_QTY As Variant, _
'                ByVal arylist_Loc As Variant, insertData_Loc As Variant)

Sub inputInList(ByRef arylist_PN As Variant, insertData_PN As Variant, _
                ByRef AryList_QTY As Variant, insertData_QTY As Variant, _
                ByRef arylist_Loc As Variant, insertData_Loc As Variant)

      
    
      
                If arylist_PN.contains(insertData_PN) Then
                    
                    indexPN = arylist_PN.indexof(insertData_PN, 0)
                    
                    If arylist_Loc(indexPN) <> "" Then
                        tempString = arylist_Loc(indexPN) & "," & insertData_Loc
                    Else
                         tempString = insertData_Loc
                    End If
                    
                    
                    tempQTY = AryList_QTY(indexPN) + insertData_QTY
                    
                    
                    arylist_Loc.removeat (indexPN)
                    arylist_Loc.Insert indexPN, tempString
                    AryList_QTY.removeat (indexPN)
                    AryList_QTY.Insert indexPN, tempQTY
                   

                Else
                    arylist_PN.Add insertData_PN
                    AryList_QTY.Add insertData_QTY
                    arylist_Loc.Add insertData_Loc
                    
                End If
        
      
      
      
End Sub

'subModule.compareLocation(data_PN, data_QTY, data_Loc, temp_title, temp_Loc, lv4_PN, lv4_QTY, lv4_Loc)
Sub compareLocation(ByRef temp_data_PN As Variant, ByRef temp_data_QTY As Variant, ByRef temp_data_Loc, _
                    ByRef top_bot_title As Variant, ByRef top_bot_Loc As Variant, _
                    ByRef temp_save_PN As Variant, ByRef temp_save_QTY As Variant, ByRef temp_save_Loc As Variant)

    Dim tempLoc, tempTitle
    Dim tempAry() As String
    Dim String1, String2
    
    
    Set tempLoc = CreateObject("system.collections.arraylist")
    Set tempTitle = CreateObject("system.collections.arraylist")
    
    
    
    
    Count = temp_data_PN.Count
    
    
    For i = 0 To Count - 1
    
        If temp_data_Loc(i) <> "" Then
            
             tempAry = Split(temp_data_Loc(i), ",")
 
 
            For j = 0 To UBound(tempAry)
       
             If tempTitle.contains(functionModule.getStr(tempAry(j))) Then
                 
                 index_titleReg = tempTitle.indexof(functionModule.getStr(tempAry(j)), 0)
                 
                 Call subModule.inputInList_Loc(tempLoc, tempAry(j), index_titleReg)
                 
                 
             Else
                 tempLoc.Add tempAry(j)
                 tempTitle.Add functionModule.getStr(tempAry(j))
             End If
            
            Next
            
            
            saveString = ""
            saveString1 = ""
            'Function compareString(ByRef String1 As Variant, ByRef String2 As Variant, ignorType As Variant) As String
            For k = 0 To tempTitle.Count - 1
                
                
                
                If top_bot_title.contains(tempTitle(k)) Then
                
                    String1 = tempLoc(k)
                    String2 = top_bot_Loc(top_bot_title.indexof(tempTitle(k), 0))
             
             
                    'For savinge
                    tempComString = compareString(String1, String2, ",")
                        
                    
                    If tempComString <> "" Then
                    
                        If saveString = "" Then
                            saveString = tempComString
                            
                        Else
                            saveString = saveString & "," & tempComString
                        End If
                    
                    End If
                    
                        
                    
                    'Modify temp_data_PN
                    If String1 <> "" Then
                    
                       If saveString1 = "" Then
                        saveString1 = String1
                        
                       Else
                        saveString1 = saveString1 & "," & String1
                       End If
                    
                    
                    
                    End If
                             
                    'overwrite dat
                    Call subModule.inputInList_overwrite(top_bot_Loc, String2, top_bot_title.indexof(tempTitle(k), 0))
                
                Else

                            If saveString1 = "" Then
                             saveString1 = tempLoc(k)
                            Else
                             saveString1 = saveString1 & "," & tempLoc(k)
                            End If

                    
                End If
                
             
             
            Next
            
             
             
             
             'modify data
             
             
             
             Call subModule.inputInList_overwrite(temp_data_Loc, saveString1, i)
             Call subModule.inputInList_overwrite(temp_data_QTY, UBound(Split(saveString1, ",")) + 1, i)
             
             
             
             'For lv45
             If saveString <> "" Then
             
                temp_save_PN.Add temp_data_PN(i)
                temp_save_QTY.Add UBound(Split(saveString, ",")) + 1
                temp_save_Loc.Add saveString
             End If
             
             
             
            
            tempTitle.Clear
            tempLoc.Clear
        
        End If
    
        'classif location title
      
       
       
    
    Next
    
    
    
    
    
    
    
    
    
    
    
    
    
End Sub


Sub combinArryList(ByRef mainList As Variant, ByRef beAppendList As Variant)

    tempCount = beAppendList.Count
    
    'Clear empty
    
    For i = 0 To tempCount - 1
    
        mainList.Add beAppendList(i)
        
    
    Next
    
    
End Sub


Sub removeEmptyAry(ByRef AryList_QTY As Variant, ByRef arylist1 As Variant, ByRef arylist2 As Variant)

    tempCount = AryList_QTY.Count
    
    'Clear empty
    
    Count = 0
    
    For i = 0 To tempCount - 1
    
        If AryList_QTY(Count) = 0 Then
        
            AryList_QTY.removeat (Count)
            arylist1.removeat (Count)
            arylist2.removeat (Count)
            
            
        Else
            Count = Count + 1
        End If
    
    
    Next
    

End Sub


Sub inputInListBox(ByRef arrylist_PN As Variant, ByRef arrylist_Loc As Variant)
    
    Dim tempAry
    
    UserForm1.ListBox1.ColumnCount = 2
    
    tempListCount = arrylist_PN.Count

    itemCount = 0

    For i = 0 To tempListCount - 1
    
        tempAry = Split(arrylist_Loc(i), ",")
        
            For j = 0 To UBound(tempAry)
                UserForm1.ListBox1.Clear
                UserForm1.ListBox1.AddItem
                UserForm1.ListBox1.List(itemCount, 0) = i 'arrylist_PN(i)
                UserForm1.ListBox1.List(itemCount, 1) = j 'tempAry(j)
                
                itemCount = itemCount + 1
            Next
    
    Next


End Sub


Sub getAryList(SheetName As Variant, start_Index As Integer, ByRef arylist1 As Variant, arylist1_index As Integer, _
                                                           ByRef arylist2 As Variant, arylist2_index As Integer, _
                                                           ByRef arylist3 As Variant, arylist3_index As Integer)

    With Workbooks(get_FileName_MakeBOM).Worksheets(SheetName)
    
        Count = 0
        
        Do While .Cells(Count + start_Index, arylist1_index).Value <> ""
        
        
            arylist1.Add .Cells(Count + start_Index, arylist1_index).Value
            arylist2.Add .Cells(Count + start_Index, arylist2_index).Value
            arylist3.Add .Cells(Count + start_Index, arylist3_index).Value
        
        
            Count = Count + 1
            
            If Count > 10000 Then
                Exit Do
            End If

        Loop
    End With


End Sub
Sub open_ExcelFile_Gene(FileName As Variant, SheetName As Variant, activeBook As Variant, sheetSel As Variant, color As Integer)
    
    Workbooks.Open FileName:=FileName
    openActive = ActiveWorkbook.Name
   
   If Sheets.Count = 1 Then
    Sheets.Add after:=Sheets(Worksheets.Count)
   End If
   
        Workbooks(openActive).Sheets(sheetSel).Move after:=Workbooks(activeBook).Sheets(Workbooks(activeBook).Worksheets.Count)
        Workbooks(activeBook).Sheets(Worksheets.Count).Name = SheetName
        Workbooks(activeBook).Sheets(Worksheets.Count).Tab.ColorIndex = color
        
        tempAry = Split(FileName, "\")
        Workbooks(tempAry(UBound(tempAry))).Close SaveChanges:=False
End Sub
Sub recovery()
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
End Sub



Sub saveSetData_col_para()

    Dim checkValue, tempCheckValue
    
    
    Set checkValue = CreateObject("system.collections.arraylist")
    Set tempCheckValue = CreateObject("system.collections.arraylist")
    
     With Workbooks(get_FileName_MakeBOM)
    
    
    
    

    tempCheckValue.Add .Worksheets("Main").Range("B34").Value
    tempCheckValue.Add .Worksheets("Main").Range("B35").Value
    tempCheckValue.Add .Worksheets("Main").Range("B36").Value
    tempCheckValue.Add .Worksheets("Main").Range("B37").Value
    tempCheckValue.Add .Worksheets("Main").Range("B38").Value
    tempCheckValue.Add .Worksheets("Main").Range("B39").Value
    tempCheckValue.Add .Worksheets("Main").Range("B40").Value
    
    checkValue.Add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    On Error GoTo exit1
    
        If checkWord(UserForm1.TextBox8) Then 'PartNumber
            Call saveData("MAIN", UserForm1.TextBox8.Value, "B34")
            
        End If
        
        If checkWord(UserForm1.TextBox9) Then 'QTY
            Call saveData("MAIN", UserForm1.TextBox9.Value, "B35")
            
        End If
        
        If checkWord(UserForm1.TextBox10) Then 'Location
            Call saveData("MAIN", UserForm1.TextBox10.Value, "B36")
        End If
        
        If checkWord(UserForm1.TextBox11) Then 'Description
            Call saveData("MAIN", UserForm1.TextBox11.Value, "B37")
        End If
        
        If checkWord(UserForm1.TextBox14) Then 'DIP and SMD
            Call saveData("MAIN", UserForm1.TextBox14.Value, "B38")
        End If
        
        If checkWord(UserForm1.TextBox16) Then 'Component type
            Call saveData("MAIN", UserForm1.TextBox16.Value, "B39")
        End If
        
        functionModule.updateUserFormValue 5
    
            
            
        If IsNumeric(UserForm1.TextBox15.Value) Then
            If Int(UserForm1.TextBox15.Value) / UserForm1.TextBox15.Value = 1 Then
            Call subModule.saveData("Main", UserForm1.TextBox15.Value, "B40")
            functionModule.updateUserFormValue 6
            End If
        Else
            GoTo exit1
        End If
        
        
        
     tempCheckValue.Add .Worksheets("Main").Range("B34").Value
    tempCheckValue.Add .Worksheets("Main").Range("B35").Value
    tempCheckValue.Add .Worksheets("Main").Range("B36").Value
    tempCheckValue.Add .Worksheets("Main").Range("B37").Value
    tempCheckValue.Add .Worksheets("Main").Range("B38").Value
    tempCheckValue.Add .Worksheets("Main").Range("B39").Value
    tempCheckValue.Add .Worksheets("Main").Range("B40").Value
    
    checkValue.Add (tempCheckValue.toArray)
    tempCheckValue.Clear
    
  
    For i = 0 To UBound(checkValue(0))
    
        If checkValue(0)(i) <> checkValue(1)(i) Then
            
            Call subModule.saveData("Main", False, "M30")
            'updateUserFormValue 7
            MsgBox "Saved!!"
            Exit For
        End If

    Next
    
    
    
    End With
 
    
     
     
exit1:

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub


Sub initProgess_1(tempUserForm As Object, procThing As String)

    With tempUserForm
    
        .progressBar.Width = 0
       
        .caption = procThing & " 0% complete"
        
    
    End With

End Sub
