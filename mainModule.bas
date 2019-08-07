Attribute VB_Name = "mainModule"
Sub load()

    Dim protectSheet
    Dim newBom, oldBom
    Dim bomArry, fileNameAry, bomSheetSel ', bomSumArry(1)
   
    
    
    
    
    
    
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With


    
    'Micro_book = Application.ActiveWorkbook.Name
    Micro_book = get_FileName_MakeBOM
    
    
    
        If UserForm1.TextBox7.Value <> "" Then
                
                
                
                
                Select Case functionModule.checkFileType(UserForm1.TextBox7.Value)
                Case 0 'open_ExcelFile(fileName As String, sheetSel As String, sheetName As String, activeBook As String)
                    Call subModule.open_ExcelFile(UserForm1.TextBox7.Value, "BOM", Micro_book, UserForm1.ComboBox1.Text)
                Case 1
                    Call subModule.open_ConceptFile(UserForm1.TextBox7.Value, "BOM", Micro_book)
                'Case 2
                    'Call subModule.open_OrcadFile
                Case Else
                
                End Select
                
                 
           
        
            
            
            
            
                
        End If




   

    
    
exit1:
    
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub
' General feature

Sub load_Gene(makeSheetName As Variant, textBoxValue As Variant, comboBoxTxt As Variant)

    Dim protectSheet
    Dim newBom, oldBom
    Dim bomArry, fileNameAry, bomSheetSel ', bomSumArry(1)
   
    
    
    
    
    
    
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With


    
    'Micro_book = Application.ActiveWorkbook.Name
    Micro_book = get_FileName_MakeBOM
    
        Dummy11 = functionModule.delSheet(makeSheetName, False, 12)
    
        If textBoxValue <> "" Then
                
                
                
                
                Select Case functionModule.checkFileType(textBoxValue)
                Case 0 'open_ExcelFile(fileName As String, sheetSel As String, sheetName As String, activeBook As String)
                    Call subModule.open_ExcelFile_Gene(textBoxValue, makeSheetName, Micro_book, comboBoxTxt, 15)
                Case 1
                    'Call subModule.open_ConceptFile(textBoxValue, makeSheetName, Micro_book)
                'Case 2
                    'Call subModule.open_OrcadFile
                Case Else
                
                End Select
                
                 
           
        
            
            
            
            
                
        End If




   

    
    
exit1:
    
     With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub


Sub get_brdLocation_csv(SheetName As String)
    
    Dim topListRegHead, topListCount, topListLocation
    Dim botListRegHead, botListCount, botListLocation
    
    
    Set topListRegHead = CreateObject("system.collections.arraylist")
    Set topListCount = CreateObject("system.collections.arraylist")
    Set topListLocation = CreateObject("system.collections.arraylist")
    
    Set botListRegHead = CreateObject("system.collections.arraylist")
    Set botListCount = CreateObject("system.collections.arraylist")
    Set botListLocation = CreateObject("system.collections.arraylist")
    
    
    
    
    
    
    
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
    
    
    
    indexstart = 1
    
    emptyCount = 10
    indexString = "REFDES"
    
    With Workbooks(get_FileName_MakeBOM).Worksheets(SheetName)
        
        
        'Get Data index
        Do While emptyCount > 0
            
            indexstart = indexstart + 1
            
            If .Cells(indexstart - 1, 1).Value = "" Then
                emptyCount = emptyCount - 1
            Else
                emptyCoun = 10
            
            End If
            
            
            
            If .Cells(indexstart - 1, 1).Value = indexString Then
                Exit Do
            End If
        Loop
        '''''''''''''''''''''''''''''''''''''''''
        
        
        Do While .Cells(indexstart, 1).Value <> ""
        
        
            If .Cells(indexstart, 5).Value Like "TP*" Then
                'filter TP item
            Else
            
                tempStr = functionModule.getStr(.Cells(indexstart, 1).Value)
                If UCase(.Cells(indexstart, 9).Value) = "YES" Then
                    
                        
                        If botListRegHead.contains(tempStr) Then
                        
                            indexNum = botListRegHead.indexof(tempStr, 0)
                        
                            countNum = botListCount(indexNum) + 1
                            tempString = botListLocation(indexNum) & "," & .Cells(indexstart, 1).Value
                            
                            botListCount.removeat (indexNum)
                            botListLocation.removeat (indexNum)
                            
                            botListCount.Insert indexNum, countNum
                            botListLocation.Insert indexNum, tempString
                            
                            
                        
                        Else
                         botListRegHead.Add tempStr
                         botListCount.Add 1
                         botListLocation.Add .Cells(indexstart, 1).Value
                        End If
           
                ElseIf UCase(.Cells(indexstart, 9).Value) = UCase("NO") Then
                        
                        If topListRegHead.contains(tempStr) Then
                        
                            indexNum = topListRegHead.indexof(tempStr, 0)
                        
                            countNum = topListCount(indexNum) + 1
                            tempString = topListLocation(indexNum) & "," & .Cells(indexstart, 1).Value
                            
                            topListCount.removeat (indexNum)
                            topListLocation.removeat (indexNum)
                            
                            topListCount.Insert indexNum, countNum
                            topListLocation.Insert indexNum, tempString
                            
                            
                        
                        Else
                         topListRegHead.Add tempStr
                         topListCount.Add 1
                         topListLocation.Add .Cells(indexstart, 1).Value
                        End If
                
                Else
                    MsgBox "Exception_not_TOP_or_BOT"
                   
                End If
            End If
            
            
        
        
            indexstart = indexstart + 1
            
            
            If indexstart = 10000 Then Exit Do
            
            
            DoEvents
        Loop
        
        
        
        
        
    
        
    End With
    
        'Sort
        If topListLocation.Count > 0 Then
            For i = 0 To topListLocation.Count - 1
                tempAry = Split(topListLocation(i), ",")
                Call Sort.QuickSort_top(tempAry, 0, UBound(tempAry))
                tempString_loc = functionModule.aryToString(tempAry)
                topListLocation(i) = tempString_loc
            Next
        Else
            MsgBox "EXCEPTION SROT"
        End If
        
        
        If botListLocation.Count > 0 Then
            For i = 0 To botListLocation.Count - 1
                tempAry = Split(botListLocation(i), ",")
                Call Sort.QuickSort_top(tempAry, 0, UBound(tempAry))
                tempString_loc = functionModule.aryToString(tempAry)
                botListLocation(i) = tempString_loc
            Next
        Else
            MsgBox "EXCEPTION SROT"
        End If
        
       
    
    
        'Modify here'''''''''''''''''''
        dataSheet = "TOP"
        Call creatSheet(dataSheet, 0, 0)
        Call subModule.printDataInSheet_txt(1, 1, dataSheet, topListRegHead)
        Call subModule.printDataInSheet_txt(1, 2, dataSheet, topListCount)
        Call subModule.printDataInSheet_txt(1, 3, dataSheet, topListLocation)
        dataSheet = "BOT"
        Call creatSheet(dataSheet, 0, 0)
        Call subModule.printDataInSheet_txt(1, 1, dataSheet, botListRegHead)
        Call subModule.printDataInSheet_txt(1, 2, dataSheet, botListCount)
        Call subModule.printDataInSheet_txt(1, 3, dataSheet, botListLocation)
    
        
        
     topListRegHead.Clear
     topListCount.Clear
     topListLocation.Clear
    
     botListRegHead.Clear
     botListCount.Clear
     botListLocation.Clear
        
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With
    
End Sub



Sub get_brdLocation(dataPath As String, dataSheet As Variant)

    Dim tempAry() As String
    Dim FileName As String
    Dim FilePath As String
    Dim tempRegHead, tempCount, tempLocation
    Dim temp() As String
    
    
    Set tempRegHead = CreateObject("system.collections.arraylist")
    Set tempCount = CreateObject("system.collections.arraylist")
    Set tempLocation = CreateObject("system.collections.arraylist")
    

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With

    Call creatSheet(dataSheet, 0, 0)


    With Workbooks(get_FileName_MakeBOM)
        
        'Load text file
        
    
        FilePath = dataPath
        
        Open FilePath For Input As #1
        
        
        j = 1
        ignor = 10
        
        
        Do Until EOF(1)
        
          Line Input #1, LineFromFile
          
          
          If j > ignor Then
            
            Do While InStr(LineFromFile, "  ") > 0
                    LineFromFile = Replace(LineFromFile, "  ", " ")
                Loop
                
                temp() = Split(LineFromFile, " ")
                   ' temp() = Split(Cells(j, 1), " ")
                
                
    
                
                If temp(2) Like "TP*" Or temp(3) Like "TP*" Then
                    
                    'Selection.Delete Shift:=xlUp
                Else
                    tempStr = functionModule.getStr(temp(0))
                    If tempRegHead.contains(tempStr) Then
                    
                        indexNum = tempRegHead.indexof(tempStr, 0)
                    
                        countNum = tempCount(indexNum) + 1
                        tempString = tempLocation(indexNum) & "," & temp(0)
                        
                        tempCount.removeat (indexNum)
                        tempLocation.removeat (indexNum)
                        
                        tempCount.Insert indexNum, countNum
                        tempLocation.Insert indexNum, tempString
                        
                        
                    
                    Else
                     tempRegHead.Add tempStr
                     tempCount.Add 1
                     tempLocation.Add temp(0)
                    End If
    
                End If
          
          End If
          
          
        
            j = j + 1
        Loop
        
        Close #1
        
        
        Call subModule.printDataInSheet_txt(1, 1, dataSheet, tempRegHead)
        Call subModule.printDataInSheet_txt(1, 2, dataSheet, tempCount)
        Call subModule.printDataInSheet_txt(1, 3, dataSheet, tempLocation)
        
        
        
    End With
    
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
    End With

End Sub


Sub get_brdLocation_backup(dataPath As String, dataSheet As Variant)

    Dim tempAry() As String
    Dim FileName As String
    Dim FilePath As String
    Dim tempRegHead
    Dim temp() As String
    
    
    Set tempRegHead = CreateObject("system.collections.arraylist")

With Workbooks(get_FileName_MakeBOM)
    
    'Load text file
    

    FilePath = dataPath
    
    Open FilePath For Input As #1
    
    j = 1
    
    Do Until EOF(1)
    
      Line Input #1, LineFromFile
    
     .Worksheets(dataSheet).Cells(j, 1).Value = LineFromFile
        j = j + 1
    Loop
    
    Close #1
    
    '''''modify
    
    .Worksheets(dataSheet).Rows("1:10").Delete Shift:=xlUp
    
    
    expCount = 1
    
        Do While .Worksheets(dataSheet).Cells(1, 1) <> ""
    
            tempString = .Worksheets(dataSheet).Cells(1, 1)
            
            Do While InStr(tempString, "  ") > 0
                tempString = Replace(tempString, "  ", " ")
            Loop
            
            temp() = Split(tempString, " ")
               ' temp() = Split(Cells(j, 1), " ")
            
            

            
            If temp(2) Like "TP*" Or temp(3) Like "TP*" Then
                
                'Selection.Delete Shift:=xlUp
            Else
                tempStr = functionModule.getStr(temp(0))
                If tempRegHead.contains(tempStr) Then
                Else
                 tempRegHead.Add tempStr
                End If

            End If
            
            
            
            
            .Worksheets(dataSheet).Rows(1).Delete Shift:=xlUp
            
            expCount = expCount + 1
            
            If expCount > 20000 Then
                Exit Do
            End If
            
        Loop
    
        Stop
    
    
    
End With

End Sub


Sub classifi(ByRef data_PN As Variant, ByRef data_QTY As Variant, ByRef data_Loc As Variant)

    Dim loc_PN, loc_QTY, loc_Loc, loc_Desc, loc_DipSmd, loc_Type, loc_DataRow
    Dim lv3_PN, lv3_QTY, lv3_Loc
    Dim lv4_PN, lv4_QTY, lv4_Loc
    Dim lv5_PN, lv5_QTY, lv5_Loc
    Dim check_PN, check_QTY, check_Loc
    'Dim data_PN, data_QTY, data_Loc
    Dim temp_PN, temp_QTY, temp_Loc
    Dim inputdata, indexPN, inputQTY
    Dim tempAry
    Dim tempList

    'init
    
    'lv3_PartNunber = UserForm1.TextBox1
    'lv4_PartNunber = UserForm1.TextBox2
    'lv5_PartNunber = UserForm1.TextBox3
    
    lv3_PartNunber = "31XXXXXXXXX"
    lv4_PartNunber = "41XXXXXXXXX"
    lv5_PartNunber = "51XXXXXXXXX"
    pcb_PartNunber = "DXXXXXXXXXX"
    
    loc_PN = UserForm1.TextBox20.Value
    loc_QTY = UserForm1.TextBox21.Value
    loc_Loc = UserForm1.TextBox22.Value
    loc_Desc = UserForm1.TextBox24.Value
    loc_DipSmd = UserForm1.TextBox23.Value
    loc_Type = UserForm1.TextBox25.Value
    loc_DataRow = UserForm1.TextBox19.Value
    
    title_lable = UserForm1.TextBox17.Value
    
    
    Set lv3_PN = CreateObject("system.collections.arraylist")
    Set lv3_QTY = CreateObject("system.collections.arraylist")
    Set lv3_Loc = CreateObject("system.collections.arraylist")
    
    Set lv4_PN = CreateObject("system.collections.arraylist")
    Set lv4_QTY = CreateObject("system.collections.arraylist")
    Set lv4_Loc = CreateObject("system.collections.arraylist")
    
    Set lv5_PN = CreateObject("system.collections.arraylist")
    Set lv5_QTY = CreateObject("system.collections.arraylist")
    Set lv5_Loc = CreateObject("system.collections.arraylist")
    
    
    Set check_PN = CreateObject("system.collections.arraylist")
    Set check_QTY = CreateObject("system.collections.arraylist")
    Set check_Loc = CreateObject("system.collections.arraylist")
    
    'Set data_PN = CreateObject("system.collections.arraylist")
    'Set data_QTY = CreateObject("system.collections.arraylist")
    'Set data_Loc = CreateObject("system.collections.arraylist")
    
    Set temp_title = CreateObject("system.collections.arraylist")
    Set temp_QTY = CreateObject("system.collections.arraylist")
    Set temp_Loc = CreateObject("system.collections.arraylist")
    
    Set tempList = CreateObject("system.collections.arraylist")
    
    With Workbooks(get_FileName_MakeBOM)
    
        Do While Replace(.Worksheets("BOM").Cells(loc_DataRow, loc_PN).Value, " ", "") <> "" Or Replace(.Worksheets("BOM").Cells(loc_DataRow, loc_Type).Value, " ", "") <> ""
        
            
             
            If functionModule.IsLike(.Worksheets("BOM").Cells(loc_DataRow, loc_PN).Value, title_lable) Or _
            (.Worksheets("BOM").Cells(loc_DataRow, loc_DipSmd).Value = "DIP" And .Worksheets("BOM").Cells(loc_DataRow, loc_Desc).Value Like "*DIP*") Or _
            .Worksheets("BOM").Cells(loc_DataRow, loc_Type).Value = "DIP" Or _
            .Worksheets("BOM").Cells(loc_DataRow, loc_Type).Value = "THMT" Then 'Find Label and DIP    through-hole mounting technology
                
                Call subModule.inputInList(lv3_PN, .Worksheets("BOM").Cells(loc_DataRow, loc_PN).Value, _
                                           lv3_QTY, .Worksheets("BOM").Cells(loc_DataRow, loc_QTY).Value, _
                                           lv3_Loc, .Worksheets("BOM").Cells(loc_DataRow, loc_Loc).Value)
                
            Else
            
             If functionModule.IsLike(.Worksheets("BOM").Cells(loc_DataRow, loc_Desc).Value, "*DIP*") Then
                
                Call subModule.inputInList(check_PN, .Worksheets("BOM").Cells(loc_DataRow, loc_PN).Value, _
                                           check_QTY, .Worksheets("BOM").Cells(loc_DataRow, loc_QTY).Value, _
                                           check_Loc, .Worksheets("BOM").Cells(loc_DataRow, loc_Loc).Value)
             Else
                Call subModule.inputInList(data_PN, .Worksheets("BOM").Cells(loc_DataRow, loc_PN).Value, _
                                           data_QTY, .Worksheets("BOM").Cells(loc_DataRow, loc_QTY).Value, _
                                           data_Loc, .Worksheets("BOM").Cells(loc_DataRow, loc_Loc).Value)
             End If
            
            End If
        
        
            loc_DataRow = loc_DataRow + 1
            
            'break
            If loc_DataRow > 10000 Then Exit Do
            
            
            DoEvents
            
        Loop
        
        
        
        'Call subModule.creatSheet_lv(lv3_PartNunber, True, 29)
        'Call subModule.printDataInSheet_lv(2, lv3_PartNunber, lv3_PN, 2, lv3_QTY, 6, lv3_Loc, 7)
        
        lvCount = 3
        
        
        
        'If UserForm1.TextBox5.Enabled Then templist.Add "TOP"
    
        'If UserForm1.TextBox6.Enabled Then templist.Add "BOT"
        
        
        'tempAry = templist.toarray()
        
        tempAry = Array("TOP", "BOT")
        For i = 0 To UBound(tempAry)
        
            With .Worksheets(tempAry(i))  'get Top/Bot Item
                
                Count = 1
                
                Do While .Cells(Count, 1).Value <> 0
                
                    temp_title.Add .Cells(Count, 1).Value
                    temp_QTY.Add .Cells(Count, 2).Value
                    temp_Loc.Add .Cells(Count, 3).Value
                    
                    Count = Count + 1
                    
                    If Count > 10000 Then GoTo exit1:
                Loop
                
                
            End With
            
                
                Select Case i
                
                Case 0
                
                    Call subModule.compareLocation(data_PN, data_QTY, data_Loc, temp_title, temp_Loc, lv4_PN, lv4_QTY, lv4_Loc)
                    
                    If lv4_PN.Count <> 0 Then
                        'Call subModule.creatSheet_General(lv4_PartNunber, True, 29, _
                                                            Array("Parent", "Part Number", "Item Number", "Alt Grp", "Usage(%)", "Qty", "Location"))
                        'Call subModule.printDataInSheet_lv(2, lv4_PartNunber, lv4_PN, 2, lv4_QTY, 6, lv4_Loc, 7)
                        lvCount = lvCount + 1
                    End If
                Case 1
                    Call subModule.compareLocation(data_PN, data_QTY, data_Loc, temp_title, temp_Loc, lv5_PN, lv5_QTY, lv5_Loc)
                    
                    If lv5_PN.Count <> 0 Then
                        'Call subModule.creatSheet_General(lv5_PartNunber, True, 29, _
                                                            Array("Parent", "Part Number", "Item Number", "Alt Grp", "Usage(%)", "Qty", "Location"))
                        'Call subModule.printDataInSheet_lv(2, lv5_PartNunber, lv5_PN, 2, lv5_QTY, 6, lv5_Loc, 7)
                        lvCount = lvCount + 1
                        
                        If lvCount = 4 Then
                            lv4_PartNunber = lv5_PartNunber
                            lv4_PN = lv5_PN
                            lv4_QTY = lv5_QTY
                            lv4_Loc = lv5_Loc
                        End If
                        
                    End If
                End Select
                
                
              
            temp_title.Clear
            temp_QTY.Clear
            temp_Loc.Clear
            
            Call subModule.saveData("MAIN", lvCount, "J22")

            
        Next
        
        
        'lv 5
        If lvCount = 5 Then
        
            lv4_PN.Insert 0, lv5_PartNunber
            lv4_QTY.Insert 0, 1
            lv4_Loc.Insert 0, ""
            
            lv5_PN.Add pcb_PartNunber
            lv5_QTY.Add 1
            lv5_Loc.Add ""
            
            Call subModule.creatSheet_General(lv5_PartNunber, True, 29, _
                                                            Array("Parent", "Part Number", "Item Number", "Alt Grp", "Usage(%)", "Qty", "Location"))
            Call subModule.printDataInSheet_lv(2, lv5_PartNunber, lv5_PN, 2, lv5_QTY, 6, lv5_Loc, 7)
            
        End If
        
        
        'lv4
         If lvCount >= 4 Then
        
            lv3_PN.Insert 0, lv4_PartNunber
            lv3_QTY.Insert 0, 1
            lv3_Loc.Insert 0, ""
            
            
            If lvCount = 4 Then
                lv4_PN.Add pcb_PartNunber
                lv4_QTY.Add 1
                lv4_Loc.Add ""
            End If
            
            Call subModule.creatSheet_General(lv4_PartNunber, True, 29, _
                                                            Array("Parent", "Part Number", "Item Number", "Alt Grp", "Usage(%)", "Qty", "Location"))
            Call subModule.printDataInSheet_lv(2, lv4_PartNunber, lv4_PN, 2, lv4_QTY, 6, lv4_Loc, 7)
            
        End If
        
        
        
         'lv3
         If lvCount >= 3 Then
         
            If lvCount = 3 Then
                lv3_PN.Add pcb_PartNunber
                lv3_QTY.Add 1
                lv3_Loc.Add ""
            End If
            
            Call subModule.creatSheet_General(lv3_PartNunber, True, 29, _
                                                            Array("Parent", "Part Number", "Item Number", "Alt Grp", "Usage(%)", "Qty", "Location"))
            Call subModule.printDataInSheet_lv(2, lv3_PartNunber, lv3_PN, 2, lv3_QTY, 6, lv3_Loc, 7)
            
        End If
        
    
    
    
    Call subModule.removeEmptyAry(data_QTY, data_Loc, data_PN)
    
    
    Call subModule.combinArryList(data_PN, check_PN)
    Call subModule.combinArryList(data_QTY, check_QTY)
    Call subModule.combinArryList(data_Loc, check_Loc)
       
    Call subModule.creatSheet_General("NeedAssign", True, 3, Array("Part Number", "Location"))
    Call subModule.printDataInSheet_check(2, "NeedAssign", data_PN, data_Loc, data_QTY)
    
    
    
   
   
   
    
    
    
    End With
    
    
    
exit1:

End Sub

Sub createBOM()

    Dim templv3, templv4, templv5
    Dim temp_PN, temp_QTY, temp_Loc
    Dim finishIndex As Integer
    Dim tempAry_List
    Dim tempAryList_List
    Dim lvPNList_List
    
    
    Set temp_PN = CreateObject("system.collections.arraylist")
    Set temp_QTY = CreateObject("system.collections.arraylist")
    Set temp_Loc = CreateObject("system.collections.arraylist")
    
    Set tempAry_List = CreateObject("system.collections.arraylist")
    Set tempAryList_List = CreateObject("system.collections.arraylist")
    Set lvPNList_List = CreateObject("system.collections.arraylist")
    
    'init
    'templv3 = UserForm1.TextBox1.Text
    'templv4 = UserForm1.TextBox2.Text
    'templv5 = UserForm1.TextBox3.Text
    'tempPCB = UserForm1.TextBox4.Text
    pcb_PartNunber = "DXXXXXXXXXX"
    
    
    'tempAry_List.Add UserForm1.TextBox1.Text
    tempAry_List.Add "31XXXXXXXXX"
    lvPNList_List.Add UserForm1.TextBox1.Text
    tempAryList_List.Add UserForm1.ListBox3
    
    If UserForm1.TextBox2.Visible Then
        'tempAry_List.Add UserForm1.TextBox2.Text
        tempAry_List.Add "41XXXXXXXXX"
        lvPNList_List.Add UserForm1.TextBox2.Text
        tempAryList_List.Add UserForm1.ListBox4
    End If
    If UserForm1.TextBox3.Visible Then
        'tempAry_List.Add UserForm1.TextBox3.Text
        tempAry_List.Add "51XXXXXXXXX"
        lvPNList_List.Add UserForm1.TextBox3.Text
        tempAryList_List.Add UserForm1.ListBox5
    End If
    
    tempAry = tempAry_List.toArray()
    tempAryList = tempAryList_List.toArray()
    lvPNList = lvPNList_List.toArray()
    'tempAryList = Array(UserForm1.ListBox3, UserForm1.ListBox4, UserForm1.ListBox5)
    
    
    finishIndex = 2
    
    
    Call subModule.creatSheet_General(UserForm1.TextBox1.Text, True, 8, _
    Array("Parent", "Part Number", "Item Number", "Alt Grp", "Usage(%)", "Qty", "Location"))
    
    
    For i = 0 To UBound(tempAry)
        
        With Workbooks(get_FileName_MakeBOM).Worksheets(UserForm1.TextBox1.Text)
            .Range("A" & finishIndex & ":G" & finishIndex).Interior.color = RGB(192, 192, 192)
        End With
        
        
        Call subModule.getAryList(tempAry(i), 2, temp_PN, 2, temp_QTY, 6, temp_Loc, 7)
        
        With tempAryList(i)
        
            If (.ListCount <> 0) Then
                
                For j = 0 To .ListCount - 1
                    
                    Call subModule.inputInList(temp_PN, .List(j, 0), temp_QTY, 1, temp_Loc, .List(j, 1))
                    
                Next
                'Call subModule.inputInList(temp_PN
            Else
            End If
        
        End With
        
        
        
        
        'If i = UBound(tempAry) Then
        '    Call subModule.inputInList(temp_PN, tempPCB, temp_QTY, 1, temp_Loc, "")
        'Else
        '    Call subModule.inputInList(temp_PN, tempAry(i + 1), temp_QTY, 1, temp_Loc, "")
        'End If
        
        
        
        finishIndex = functionModule.printDataInSheet(finishIndex, UserForm1.TextBox1.Text, tempAry(i), temp_PN, 2, temp_QTY, 6, temp_Loc, 7)
        
        
        'Call subModule.creatSheet_General(tempAry(i), True, 29, _
                                                        Array("Parent", "Part Number", "Item Number", "Alt Grp", "Usage(%)", "Qty", "Location"))
        'Call subModule.printDataInSheet_lv(2, tempAry(i), temp_PN, 2, temp_QTY, 6, temp_Loc, 7)
        
        
        
        
        
        
        temp_PN.Clear
        temp_QTY.Clear
        temp_Loc.Clear
        
        DoEvents
    Next
    
 
    tempRang = "A2:B" & finishIndex - 1
    
    
    With Workbooks(get_FileName_MakeBOM).Worksheets(UserForm1.TextBox1.Text)
    
        For i = 0 To UBound(lvPNList)
            .Range(tempRang).Replace tempAry(i), lvPNList(i)
        Next
    
            .Range(tempRang).Replace pcb_PartNunber, UserForm1.TextBox4.Text
            
            .Cells.Font.Name = "Calibri"
            .Cells.Font.Size = 12
            .Range("A1:G" & finishIndex - 1).Borders.LineStyle = xlContinuous
            .Range("A1:G1").Font.color = RGB(255, 255, 255)
            .Range("A1:G1").Interior.color = RGB(0, 204, 255)
            .Columns("A:F").AutoFit
            .Columns("G").ColumnWidth = 20
            
    End With
    
    
    
    
    
    MsgBox ("Done")
    UserForm1.CommandButton40.Visible = True
End Sub
