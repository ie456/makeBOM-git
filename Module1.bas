Attribute VB_Name = "Module1"

Sub printOut(filname As Variant)

    Dim tempListAry
    
     With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
   
    
    
    With Workbooks(get_FileName_MakeBOM)
    
    
    Set wb2 = Application.Workbooks.Add
    
    SheetName = Workbooks(get_FileName_MakeBOM).Worksheets("MAIN").Range("B24").Value
    .Worksheets(SheetName).Copy after:=wb2.Sheets(wb2.Sheets.Count)
   
   
    wb2.Sheets(1).Delete
    
   
    wb2.SaveAs FileName:=get_WorkBookNamePath & "\" & filname & ".xls", FileFormat:=56
    'Wb2.SaveAs fileName:=wb1.Path & "\" & filname & ".xlsx", FileFormat:=51
    wb2.Close
 End With
 
 
MsgBox "Done!"
    

'EVENT no sheet select


exitEnd:


With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
End With

    Exit Sub



GoTo exitEnd


End Sub
Sub printOut_dir(filname As Variant, Path As Variant)

    Dim tempListAry
    
     With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
   
    
    
    With Workbooks(get_FileName_MakeBOM)
    
    
    Set wb2 = Application.Workbooks.Add
    
    SheetName = Workbooks(get_FileName_MakeBOM).Worksheets("MAIN").Range("B24").Value
    .Worksheets(SheetName).Copy after:=wb2.Sheets(wb2.Sheets.Count)
   
   
    wb2.Sheets(1).Delete
    
   
    wb2.SaveAs FileName:=Path & "\" & filname & ".xls", FileFormat:=56
    'Wb2.SaveAs fileName:=wb1.Path & "\" & filname & ".xlsx", FileFormat:=51
    wb2.Close
 End With
 
 
MsgBox "Done!"
    

'EVENT no sheet select


exitEnd:


With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
End With

    Exit Sub



GoTo exitEnd


End Sub
Sub printOut_lv(filname As Variant, Path As Variant)

    Dim tempListAry
    Dim tempArry
    
     With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
    End With
    
   
    
    
    With Workbooks(get_FileName_MakeBOM)
    
    SheetName = .Worksheets("MAIN").Range("B24").Value
    
    'Get lv
    bomlv = .Worksheets("MAIN").Range("J22").Value
    
    
    index = 2
    
    For i = 0 To bomlv - 3
    
        tempIndexStart = index
        
        tempLvPart = .Worksheets("MAIN").Cells(24 + i, 2).Value
        
        
         Count = 0
        Do While 1
        
            If .Worksheets(SheetName).Cells(index + Count, 1).Value <> tempLvPart Then
                index = index + Count
                Exit Do
            Else
                
                Count = Count + 1
                
            End If
            
            
            
            If Count > 5000 Then
                MsgBox "Exception : Over 5000 "
                Exit Sub
            End If
            
        Loop
        
        
    
         Set wb2 = Application.Workbooks.Add
         
         wb2.Sheets(1).Cells(1, 1).Value = "Parent"
         wb2.Sheets(1).Cells(1, 2).Value = "Part Number"
         wb2.Sheets(1).Cells(1, 3).Value = "Item Number"
         wb2.Sheets(1).Cells(1, 4).Value = "Alt Grp"
         wb2.Sheets(1).Cells(1, 5).Value = "Usage(%)"
         wb2.Sheets(1).Cells(1, 6).Value = "Qty"
         wb2.Sheets(1).Cells(1, 7).Value = "Location"
         
         
        
         .Worksheets(SheetName).Range("A" & tempIndexStart & ":G" & index - 1).Copy
         wb2.Sheets(1).Range("A2").PasteSpecial Paste:=xlPasteValues
         
        
        
        With wb2.Sheets(1)
            .Cells.Font.Name = "Calibri"
            .Cells.Font.Size = 12
            .Range("A1:G" & index - tempIndexStart + 1).Borders.LineStyle = xlContinuous
            .Range("A1:G1").Font.color = RGB(255, 255, 255)
            .Range("A1:G1").Interior.color = RGB(0, 204, 255)
            .Columns("A:F").AutoFit
            .Columns("G").ColumnWidth = 20
            
        End With
        
        
        
        
        
        
         wb2.SaveAs FileName:=Path & "\" & filname & "_" & i + 3 & ".xls", FileFormat:=56
         'Wb2.SaveAs fileName:=wb1.Path & "\" & filname & ".xlsx", FileFormat:=51
         wb2.Close
        
    Next
    
 End With
 
 
MsgBox "Done!"
    

'EVENT no sheet select


exitEnd:


With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
End With

    Exit Sub



GoTo exitEnd


End Sub
Sub remove()


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
End Sub

Sub test()
    UserForm2.show
End Sub
